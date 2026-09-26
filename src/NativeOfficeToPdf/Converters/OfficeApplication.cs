using System.Diagnostics;
using System.Runtime.InteropServices;
using NativeOfficeToPdf.Interop;

namespace NativeOfficeToPdf.Converters;

/// <summary>
/// Ciclo de vida de una aplicación de Office usada como servidor COM.
/// <para>
/// La regla que gobierna esta clase: <b>si la instancia ya estaba en marcha, es del usuario y no se
/// cierra</b>. Cerrar de golpe el Word que alguien tenía abierto con un documento sin guardar es el
/// error clásico de este tipo de herramientas. Solo se llama a <c>Quit</c> sobre las instancias que
/// levantamos nosotros.
/// </para>
/// <para>
/// Mientras dura, la instancia no muestra alertas y abre los documentos con las macros desactivadas,
/// para que un AutoOpen no se ejecute ni deje un diálogo modal esperando a nadie. Al terminar se
/// restauran los valores que tenía: si la instancia es del usuario, su configuración no cambia.
/// </para>
/// </summary>
internal sealed class OfficeApplication : IDisposable
{
    // msoAutomationSecurityForceDisable
    private const int MsoAutomationSecurityForceDisable = 3;

    private const int CoEServerExecFailure = unchecked((int)0x80080005);

    private readonly object?[] quitArguments;
    private readonly object? previousAlerts;
    private readonly object? previousSecurity;
    private bool disposed;

    private OfficeApplication(ComObject application, bool startedByUs, int alertsNone, object?[] quitArguments)
    {
        Application = application;
        StartedByUs = startedByUs;
        this.quitArguments = quitArguments;

        previousAlerts = application.TryGetProperty("DisplayAlerts");
        previousSecurity = application.TryGetProperty("AutomationSecurity");
        application.TrySetProperty("DisplayAlerts", alertsNone);
        application.TrySetProperty("AutomationSecurity", MsoAutomationSecurityForceDisable);
    }

    public ComObject Application { get; }

    /// <summary>Verdadero solo si esta instancia la creamos nosotros y por lo tanto podemos cerrarla.</summary>
    public bool StartedByUs { get; }

    /// <param name="progId">Por ejemplo <c>Word.Application</c>.</param>
    /// <param name="processName">Nombre del proceso sin extensión, por ejemplo <c>WINWORD</c>.</param>
    /// <param name="alertsNone">Valor de <c>DisplayAlerts</c> que silencia la aplicación; cada una usa el suyo.</param>
    /// <param name="quitArguments">Argumentos posicionales de <c>Quit</c>, si la aplicación los acepta.</param>
    public static OfficeApplication GetOrCreate(
        string progId, string processName, int alertsNone, params object?[] quitArguments)
    {
        object? active = TryGetActive(progId);
        if (active is not null)
        {
            return new OfficeApplication(new ComObject(active), startedByUs: false, alertsNone, quitArguments);
        }

        Type? type = Type.GetTypeFromProgID(progId, throwOnError: false);
        if (type is null)
        {
            throw new OfficeAutomationException(
                $"No se encontró el componente COM '{progId}'. ¿Está instalada esa aplicación de Office?");
        }

        object created = Create(type, progId, processName);
        return new OfficeApplication(new ComObject(created), startedByUs: true, alertsNone, quitArguments);
    }

    /// <summary>
    /// Arranca la aplicación con un reintento. Si Office se cuelga al iniciarse —típicamente por un
    /// aviso que nadie ve cuando lo lanza la automatización—, COM espera 120 segundos a que registre
    /// su fábrica de clases y se rinde con <c>CO_E_SERVER_EXEC_FAILURE</c> (evento DCOM 10010). En ese
    /// caso se cierra el proceso colgado, que si no bloquearía las conversiones siguientes, y se
    /// reintenta una vez, como recomienda Microsoft para este error.
    /// </summary>
    private static object Create(Type type, string progId, string processName)
    {
        for (int attempt = 1; ; attempt++)
        {
            HashSet<int> before = ProcessIds(processName);
            try
            {
                return Activator.CreateInstance(type)
                    ?? throw new OfficeAutomationException($"No se pudo iniciar '{progId}'.");
            }
            catch (COMException ex) when (ex.HResult == CoEServerExecFailure)
            {
                // ponytail: "nuevo desde antes del arranque" = "lanzado por nosotros". Si el usuario abre
                // esa misma aplicación a mano durante la espera de 120 s, también se cerraría.
                KillNewProcesses(processName, before);

                if (attempt == 2)
                {
                    string app = progId.Split('.')[0];
                    throw new OfficeAutomationException(
                        $"{app} no respondió al iniciarse automáticamente para proceder con la conversión. " +
                        "Abra el documento manualmente para revisar si hay notificaciones pendientes.", ex);
                }
            }
            catch (COMException ex)
            {
                throw new OfficeAutomationException($"No se pudo iniciar '{progId}': {ex.Message}", ex);
            }
        }
    }

    private static HashSet<int> ProcessIds(string processName)
    {
        Process[] processes = Process.GetProcessesByName(processName);
        HashSet<int> ids = processes.Select(p => p.Id).ToHashSet();
        foreach (Process process in processes)
        {
            process.Dispose();
        }

        return ids;
    }

    private static void KillNewProcesses(string processName, HashSet<int> before)
    {
        foreach (Process process in Process.GetProcessesByName(processName))
        {
            using (process)
            {
                if (before.Contains(process.Id))
                {
                    continue;
                }

                try
                {
                    process.Kill();
                    process.WaitForExit(10_000);
                }
                catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
                {
                    // Ya terminó solo, o no hay permiso: el reintento dirá si sigue bloqueando.
                }
            }
        }
    }

    /// <summary>
    /// Busca en el Running Object Table la instancia que el usuario ya tiene abierta, en vez de levantar
    /// una paralela — y, sobre todo, para saber que esa instancia no es nuestra y no debemos cerrarla.
    /// Cualquier HRESULT distinto de cero se trata como "no hay instancia": el caso normal
    /// (MK_E_UNAVAILABLE) es indistinguible en la práctica de un fallo al consultar, y en ambos la
    /// respuesta correcta es crear una instancia nueva.
    /// </summary>
    private static object? TryGetActive(string progId)
    {
        if (NativeMethods.CLSIDFromProgID(progId, out Guid clsid) != 0)
        {
            return null;
        }

        return NativeMethods.GetActiveObject(ref clsid, IntPtr.Zero, out object instance) == 0
            ? instance
            : null;
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;

        if (previousSecurity is not null)
        {
            Application.TrySetProperty("AutomationSecurity", previousSecurity);
        }

        if (previousAlerts is not null)
        {
            Application.TrySetProperty("DisplayAlerts", previousAlerts);
        }

        try
        {
            if (StartedByUs)
            {
                Application.Invoke("Quit", quitArguments);
            }
        }
        catch (COMException)
        {
            // Office puede haberse cerrado solo. No es motivo para fallar una conversión ya escrita.
        }
        finally
        {
            Application.Dispose();

            // Sin esto, los RCW pendientes mantienen viva la referencia y quedan procesos WINWORD.EXE
            // o POWERPNT.EXE huérfanos. Es la causa habitual de que estas herramientas "se degraden"
            // con el uso repetido.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
