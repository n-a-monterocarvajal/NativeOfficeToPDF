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
/// </summary>
internal sealed class OfficeApplication : IDisposable
{
    private readonly object?[] quitArguments;
    private bool disposed;

    private OfficeApplication(ComObject application, bool startedByUs, object?[] quitArguments)
    {
        Application = application;
        StartedByUs = startedByUs;
        this.quitArguments = quitArguments;
    }

    public ComObject Application { get; }

    /// <summary>Verdadero solo si esta instancia la creamos nosotros y por lo tanto podemos cerrarla.</summary>
    public bool StartedByUs { get; }

    /// <param name="progId">Por ejemplo <c>Word.Application</c>.</param>
    /// <param name="quitArguments">Argumentos posicionales de <c>Quit</c>, si la aplicación los acepta.</param>
    public static OfficeApplication GetOrCreate(string progId, params object?[] quitArguments)
    {
        object? active = ActiveObject.TryGet(progId);
        if (active is not null)
        {
            return new OfficeApplication(new ComObject(active), startedByUs: false, quitArguments);
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
            return new OfficeApplication(new ComObject(created), startedByUs: true, quitArguments);
        }
        catch (COMException ex)
        {
            throw new OfficeAutomationException($"No se pudo iniciar '{progId}': {ex.Message}", ex);
        }
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;

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
