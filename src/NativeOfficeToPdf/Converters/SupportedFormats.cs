namespace NativeOfficeToPdf.Converters;

/// <summary>Aplicación de Office que sabe abrir un formato dado.</summary>
internal enum OfficeFamily
{
    Word,
    PowerPoint,
}

/// <summary>
/// Única lista de extensiones soportadas del proyecto. La consultan el despachador de conversión, el
/// registrador del menú contextual y las pruebas, de modo que agregar un formato sea un solo cambio.
/// El instalador de Inno Setup repite estas extensiones en su sección [Registry]; hay una prueba que
/// compara ambas listas para que no se desincronicen.
/// </summary>
internal static class SupportedFormats
{
    private static readonly Dictionary<string, OfficeFamily> ByExtension =
        new(StringComparer.OrdinalIgnoreCase)
        {
            [".doc"] = OfficeFamily.Word,
            [".docx"] = OfficeFamily.Word,
            [".docm"] = OfficeFamily.Word,
            [".ppt"] = OfficeFamily.PowerPoint,
            [".pptx"] = OfficeFamily.PowerPoint,
            [".pptm"] = OfficeFamily.PowerPoint,
        };

    /// <summary>Extensiones soportadas, con punto y en minúsculas.</summary>
    public static IReadOnlyCollection<string> Extensions => ByExtension.Keys;

    public static bool IsSupported(string path) => TryGetFamily(path, out _);

    public static bool TryGetFamily(string path, out OfficeFamily family)
    {
        string extension = Path.GetExtension(path);
        if (string.IsNullOrEmpty(extension))
        {
            family = default;
            return false;
        }

        return ByExtension.TryGetValue(extension, out family);
    }
}
