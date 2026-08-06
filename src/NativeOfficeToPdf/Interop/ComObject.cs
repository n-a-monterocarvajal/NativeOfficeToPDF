using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace NativeOfficeToPdf.Interop;

/// <summary>
/// Envoltorio delgado sobre un objeto COM por enlace tardío (IDispatch).
/// <para>
/// Dos razones para que exista: dar a las llamadas la forma de argumentos con nombre que tiene la API
/// de Office (<c>ExportAsFixedFormat</c> tiene catorce parámetros, casi todos opcionales, y ordenarlos
/// a mano es la fuente de error más probable de todo el proyecto), y centralizar la liberación con
/// <see cref="Marshal.ReleaseComObject"/>, que es lo que evita dejar procesos WINWORD.EXE huérfanos.
/// </para>
/// </summary>
internal sealed class ComObject : IDisposable
{
    private object? instance;

    public ComObject(object instance)
    {
        this.instance = instance ?? throw new ArgumentNullException(nameof(instance));
    }

    private object Instance =>
        instance ?? throw new ObjectDisposedException(nameof(ComObject));

    /// <summary>Llama a un método y descarta el resultado.</summary>
    public void Invoke(string name, params object?[] arguments) =>
        InvokeCore(BindingFlags.InvokeMethod, name, arguments, namedArguments: null);

    /// <summary>
    /// Llama a un método pasando los argumentos por nombre, tal como se documentan en la referencia
    /// VBA de Office. El orden del arreglo no importa mientras nombres y valores se correspondan.
    /// </summary>
    public void InvokeNamed(string name, params (string Name, object? Value)[] arguments)
    {
        object?[] values = new object?[arguments.Length];
        string[] names = new string[arguments.Length];
        for (int i = 0; i < arguments.Length; i++)
        {
            values[i] = arguments[i].Value;
            names[i] = arguments[i].Name;
        }

        InvokeCore(BindingFlags.InvokeMethod, name, values, names);
    }

    /// <summary>Llama a un método que devuelve otro objeto COM.</summary>
    public ComObject InvokeNamedForObject(string name, params (string Name, object? Value)[] arguments)
    {
        object?[] values = new object?[arguments.Length];
        string[] names = new string[arguments.Length];
        for (int i = 0; i < arguments.Length; i++)
        {
            values[i] = arguments[i].Value;
            names[i] = arguments[i].Name;
        }

        object? result = InvokeCore(BindingFlags.InvokeMethod, name, values, names);
        return Wrap(name, result);
    }

    /// <summary>Lee una propiedad escalar.</summary>
    public object? GetProperty(string name) =>
        InvokeCore(BindingFlags.GetProperty, name, Array.Empty<object?>(), namedArguments: null);

    /// <summary>Lee una propiedad que devuelve otro objeto COM (por ejemplo <c>Application.Documents</c>).</summary>
    public ComObject GetObject(string name) =>
        Wrap(name, GetProperty(name));

    /// <summary>Escribe una propiedad.</summary>
    public void SetProperty(string name, object? value) =>
        InvokeCore(BindingFlags.SetProperty, name, [value], namedArguments: null);

    /// <summary>
    /// Escribe una propiedad ignorando el fallo. Se usa solo para ajustes cosméticos sobre una
    /// instancia de Office que puede no ser nuestra y que puede rechazar el cambio según su estado.
    /// </summary>
    public void TrySetProperty(string name, object? value)
    {
        try
        {
            SetProperty(name, value);
        }
        catch (COMException)
        {
            // Ajuste opcional: si Office lo rechaza, la conversión sigue siendo válida.
        }
        catch (TargetInvocationException)
        {
        }
    }

    /// <summary>Lee una propiedad, devolviendo <paramref name="fallback"/> si Office la rechaza.</summary>
    public object? TryGetProperty(string name, object? fallback = null)
    {
        try
        {
            return GetProperty(name);
        }
        catch (COMException)
        {
            return fallback;
        }
        catch (TargetInvocationException)
        {
            return fallback;
        }
    }

    private object? InvokeCore(BindingFlags flags, string name, object?[] arguments, string[]? namedArguments)
    {
        try
        {
            return Instance.GetType().InvokeMember(
                name,
                flags,
                binder: null,
                target: Instance,
                args: arguments,
                modifiers: null,
                culture: CultureInfo.InvariantCulture,
                namedParameters: namedArguments);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            // La excepción real de Office viaja adentro; propagarla tal cual hace que el manejo de
            // COMException en Program.cs funcione y que el mensaje al usuario sea el de Office.
            throw ex.InnerException;
        }
    }

    private static ComObject Wrap(string name, object? result) =>
        result is null
            ? throw new InvalidOperationException($"Office devolvió un valor nulo para '{name}'.")
            : new ComObject(result);

    public void Dispose()
    {
        object? released = Interlocked.Exchange(ref instance, null);
        if (released is null)
        {
            return;
        }

        if (Marshal.IsComObject(released))
        {
            Marshal.ReleaseComObject(released);
        }
    }
}
