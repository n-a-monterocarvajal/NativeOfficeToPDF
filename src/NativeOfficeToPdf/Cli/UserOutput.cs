using NativeOfficeToPdf.Interop;

namespace NativeOfficeToPdf.Cli;

/// <summary>
/// El puente entre ser un CLI y ser una app sin consola.
/// <para>
/// El binario se compila como <c>WinExe</c> para que el menú contextual no haga parpadear una ventana
/// negra. Eso, por sí solo, dejaría a la herramienta muda desde la línea de comandos, así que al
/// arrancar se intenta enganchar la consola del proceso padre: si el usuario lo invocó desde
/// PowerShell o cmd, escribe ahí como cualquier CLI; si lo invocó Explorer, no hay consola y los
/// errores salen por un cuadro de diálogo. El éxito nunca muestra nada.
/// </para>
/// </summary>
internal sealed class UserOutput
{
    private const string Caption = "Convertir a PDF";

    /// <summary>
    /// Válvula de escape para uso desatendido: con <c>NATIVEOFFICETOPDF_NO_DIALOGS=1</c> la herramienta
    /// no abre ningún cuadro de diálogo aunque no tenga consola. La usan el CI y cualquier despliegue
    /// automatizado, donde un diálogo modal sería un cuelgue.
    /// </summary>
    private const string NoDialogsVariable = "NATIVEOFFICETOPDF_NO_DIALOGS";

    private readonly bool dialogsAllowed;

    public UserOutput()
    {
        HasConsole = NativeMethods.GetConsoleWindow() != IntPtr.Zero ||
                     NativeMethods.AttachConsole(NativeMethods.AttachParentProcess);

        string? suppress = Environment.GetEnvironmentVariable(NoDialogsVariable);
        dialogsAllowed = !string.Equals(suppress, "1", StringComparison.Ordinal) &&
                         !string.Equals(suppress, "true", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Verdadero si hay una consola donde escribir; falso si la invocación vino de Explorer.</summary>
    public bool HasConsole { get; }

    public void Info(string message)
    {
        if (HasConsole)
        {
            Console.Out.WriteLine(message);
        }
    }

    /// <summary>Mensaje informativo que sí debe verse aunque no haya consola.</summary>
    public void Notice(string message)
    {
        if (HasConsole || !dialogsAllowed)
        {
            Console.Out.WriteLine(message);
            return;
        }

        NativeMethods.MessageBoxW(
            IntPtr.Zero, message, Caption,
            NativeMethods.MbOk | NativeMethods.MbIconInformation | NativeMethods.MbSetForeground);
    }

    public void Error(string message)
    {
        if (HasConsole || !dialogsAllowed)
        {
            Console.Error.WriteLine(message);
            return;
        }

        NativeMethods.MessageBoxW(
            IntPtr.Zero, message, Caption,
            NativeMethods.MbOk | NativeMethods.MbIconError | NativeMethods.MbSetForeground);
    }

    /// <summary>
    /// Pregunta sí/no. Sin consola se pregunta con un cuadro de diálogo; con consola no se pregunta
    /// nada — un CLI no debe quedarse esperando en un script — y se devuelve
    /// <paramref name="consoleAnswer"/>.
    /// </summary>
    public bool Confirm(string question, bool consoleAnswer)
    {
        if (HasConsole || !dialogsAllowed)
        {
            return consoleAnswer;
        }

        int result = NativeMethods.MessageBoxW(
            IntPtr.Zero, question, Caption,
            NativeMethods.MbYesNo | NativeMethods.MbIconQuestion | NativeMethods.MbSetForeground);

        return result == NativeMethods.IdYes;
    }
}
