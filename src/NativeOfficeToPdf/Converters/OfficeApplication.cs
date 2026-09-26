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
    /// <param name="alertsNone">Valor de <c>DisplayAlerts</c> que silencia la aplicación; cada una usa el suyo.</param>
    /// <param name="quitArguments">Argumentos posicionales de <c>Quit</c>, si la aplicación los acepta.</param>
    public static OfficeApplication GetOrCreate(string progId, int alertsNone, params object?[] quitArguments)
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

        try
        {
            object created = Activator.CreateInstance(type)
                ?? throw new OfficeAutomationException($"No se pudo iniciar '{progId}'.");
            return new OfficeApplication(new ComObject(created), startedByUs: true, alertsNone, quitArguments);
        }
        catch (COMException ex)
        {
            throw new OfficeAutomationException($"No se pudo iniciar '{progId}': {ex.Message}", ex);
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
