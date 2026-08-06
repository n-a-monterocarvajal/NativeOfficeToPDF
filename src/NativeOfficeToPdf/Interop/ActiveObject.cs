namespace NativeOfficeToPdf.Interop;

/// <summary>
/// Acceso al Running Object Table. Permite reutilizar el Word o el PowerPoint que el usuario ya tiene
/// abierto, en vez de levantar una instancia paralela — y, sobre todo, permite saber que esa instancia
/// no es nuestra y por lo tanto no debemos cerrarla al terminar.
/// </summary>
internal static class ActiveObject
{
    /// <summary>
    /// Devuelve la instancia registrada para <paramref name="progId"/>, o <c>null</c> si no hay
    /// ninguna en marcha. Cualquier HRESULT distinto de cero se trata como "no hay instancia": el
    /// caso normal (MK_E_UNAVAILABLE) es indistinguible en la práctica de un fallo al consultar, y en
    /// ambos la respuesta correcta es crear una instancia nueva.
    /// </summary>
    public static object? TryGet(string progId)
    {
        if (NativeMethods.CLSIDFromProgID(progId, out Guid clsid) != 0)
        {
            return null;
        }

        return NativeMethods.GetActiveObject(ref clsid, IntPtr.Zero, out object instance) == 0
            ? instance
            : null;
    }
}
