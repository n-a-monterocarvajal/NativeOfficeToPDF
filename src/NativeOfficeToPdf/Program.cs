using System.Diagnostics;
using System.Runtime.InteropServices;
using NativeOfficeToPdf.Cli;
using NativeOfficeToPdf.Converters;
using NativeOfficeToPdf.Shell;
using NativeOfficeToPdf.Update;

namespace NativeOfficeToPdf;

internal static class Program
{
    // La automatización COM se hace desde un apartamento STA: es lo que esperan los servidores de
    // Office y evita rarezas de marshalling en las llamadas por IDispatch.
    [STAThread]
    private static int Main(string[] args)
    {
        var output = new UserOutput();
        ParsedCommand command = ArgumentParser.Parse(args);

        if (!command.IsValid)
        {
            output.Error($"{command.Error}{Environment.NewLine}{Environment.NewLine}{HelpText}");
            return ExitCodes.InvalidArguments;
        }

        try
        {
            return command.Verb switch
            {
                Verb.Convert => Convert(command, output),
                Verb.Install => Install(output),
                Verb.Uninstall => Uninstall(output),
                Verb.CheckUpdates => CheckUpdates(output),
                Verb.Version => Version(output),
                _ => Help(output),
            };
        }
        catch (OfficeAutomationException ex)
        {
            output.Error(ex.Message);
            return ExitCodes.OfficeAutomationError;
        }
        catch (COMException ex)
        {
            output.Error($"Error de automatización de Office: {ex.Message}");
            return ExitCodes.OfficeAutomationError;
        }
        catch (Exception ex)
        {
            output.Error($"Error inesperado: {ex.Message}");
            return ExitCodes.Failure;
        }
    }

    private static int Convert(ParsedCommand command, UserOutput output)
    {
        string source;
        try
        {
            source = Path.GetFullPath(command.Source!);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            output.Error($"Ruta de origen inválida: {command.Source}");
            return ExitCodes.InvalidArguments;
        }

        if (!File.Exists(source))
        {
            output.Error($"No se encontró el archivo: {source}");
            return ExitCodes.SourceNotFound;
        }

        if (!SupportedFormats.TryGetFamily(source, out OfficeFamily family))
        {
            output.Error(
                $"Extensión no soportada: {Path.GetExtension(source)}. " +
                $"Soportadas: {string.Join(", ", SupportedFormats.Extensions.Order(StringComparer.Ordinal))}.");
            return ExitCodes.UnsupportedExtension;
        }

        string destination = ConversionPaths.ResolveDestination(source, command.Destination);

        if (File.Exists(destination) && !command.Overwrite)
        {
            // Desde la consola no se pregunta nada (rompería cualquier script): se exige --overwrite.
            // Desde Explorer, donde el usuario está mirando, se pregunta.
            bool proceed = output.Confirm(
                $"«{Path.GetFileName(destination)}» ya existe.{Environment.NewLine}¿Desea reemplazarlo?",
                consoleAnswer: false);

            if (!proceed)
            {
                // Solo por consola: desde Explorer el usuario acaba de responder que no, y un cuadro
                // de error después de su propia decisión sobra —y encima hablaría de una opción del
                // CLI que ahí no puede escribir—.
                output.ConsoleError($"El destino ya existe: {destination}. Use --overwrite para reemplazarlo.");
                return ExitCodes.Failure;
            }
        }

        string? destinationDirectory = Path.GetDirectoryName(destination);
        if (!string.IsNullOrEmpty(destinationDirectory) && !Directory.Exists(destinationDirectory))
        {
            Directory.CreateDirectory(destinationDirectory);
        }

        IDocumentConverter converter = ConverterFactory.Create(family);
        converter.Convert(source, destination);

        if (!File.Exists(destination))
        {
            output.Error($"Office terminó sin errores pero no se generó el PDF: {destination}");
            return ExitCodes.Failure;
        }

        output.Info($"PDF generado: {destination}");

        if (command.Open)
        {
            Process.Start(new ProcessStartInfo(destination) { UseShellExecute = true })?.Dispose();
        }

        if (!command.Quiet)
        {
            NotifyIfUpdateAvailable(output);
        }

        return ExitCodes.Success;
    }

    private static int Install(UserOutput output)
    {
        ContextMenuRegistrar.Install(AppInfo.ExecutablePath);
        output.Notice(
            $"«{ContextMenuRegistrar.VerbDisplayName}» quedó registrado para el usuario actual, " +
            $"apuntando a:{Environment.NewLine}{AppInfo.ExecutablePath}");
        return ExitCodes.Success;
    }

    private static int Uninstall(UserOutput output)
    {
        ContextMenuRegistrar.Uninstall();
        output.Notice($"«{ContextMenuRegistrar.VerbDisplayName}» quedó fuera del menú contextual.");
        return ExitCodes.Success;
    }

    private static int CheckUpdates(UserOutput output)
    {
        UpdateChecker.Result result = UpdateChecker.Check(force: true);

        if (result.Release is null)
        {
            output.Notice(
                "No se pudo consultar si hay versiones nuevas. " +
                "Puede ser falta de red, o que el repositorio no publique Releases todavía.");
            return ExitCodes.Success;
        }

        if (!result.IsNewer)
        {
            output.Notice($"Está al día: versión {AppInfo.Version}.");
            return ExitCodes.Success;
        }

        AnnounceUpdate(output, result.Release);
        return ExitCodes.UpdateAvailable;
    }

    private static int Version(UserOutput output)
    {
        output.Notice($"{AppInfo.Name} {AppInfo.Version}");
        return ExitCodes.Success;
    }

    private static int Help(UserOutput output)
    {
        output.Notice(HelpText);
        return ExitCodes.Success;
    }

    /// <summary>
    /// Aviso pasivo después de una conversión: como mucho uno por versión nueva, y nunca antes de que
    /// el PDF esté escrito.
    /// </summary>
    private static void NotifyIfUpdateAvailable(UserOutput output)
    {
        UpdateChecker.Result result = UpdateChecker.Check(force: false);
        if (result.Release is null || !result.IsNewer || result.AlreadyNotified)
        {
            return;
        }

        AnnounceUpdate(output, result.Release);
        UpdateChecker.MarkNotified(result.Release);
    }

    private static void AnnounceUpdate(UserOutput output, ReleaseInfo release)
    {
        if (output.HasConsole)
        {
            output.Notice(
                $"Hay una versión más nueva: {release.Version} (instalada: {AppInfo.Version}).{Environment.NewLine}{release.Url}");
            return;
        }

        bool open = output.Confirm(
            $"Hay una versión más nueva de {AppInfo.Name}: {release.Version}." +
            $"{Environment.NewLine}Instalada: {AppInfo.Version}." +
            $"{Environment.NewLine}{Environment.NewLine}¿Abrir la página de descarga?",
            consoleAnswer: false);

        if (open)
        {
            Process.Start(new ProcessStartInfo(release.Url) { UseShellExecute = true })?.Dispose();
        }
    }

    private static string HelpText =>
        $"""
        {AppInfo.Name} {AppInfo.Version} — Word y PowerPoint a PDF por la vía nativa de Office.

        Uso:
          NativeOfficeToPdf.exe <origen> [destino] [opciones]
          NativeOfficeToPdf.exe convert <origen> [destino] [opciones]
          NativeOfficeToPdf.exe install | uninstall
          NativeOfficeToPdf.exe check-updates
          NativeOfficeToPdf.exe --version | --help

        Opciones:
          --overwrite, -o   Reemplaza el PDF de destino si ya existe.
          --open            Abre el PDF al terminar.
          --quiet, -q       No consulta si hay versiones nuevas.

        Extensiones soportadas:
          {string.Join(" ", SupportedFormats.Extensions.Order(StringComparer.Ordinal))}

        Códigos de salida:
          0 éxito · 1 fallo genérico · 2 argumentos inválidos · 3 extensión no soportada
          4 origen no encontrado · 5 error de Office · 6 hay actualización

        Nota para scripts: el binario es una app de Windows sin consola propia, así que PowerShell no
        espera a que termine. Use «Start-Process -Wait -PassThru» para leer el código de salida.
        """;
}
