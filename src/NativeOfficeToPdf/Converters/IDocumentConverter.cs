namespace NativeOfficeToPdf.Converters;

internal interface IDocumentConverter
{
    /// <summary>
    /// Exporta <paramref name="sourcePath"/> a <paramref name="destinationPath"/> usando el motor
    /// nativo de Office. Ambas rutas deben venir absolutas: Office resuelve las relativas contra su
    /// propio directorio de trabajo, no contra el nuestro.
    /// </summary>
    void Convert(string sourcePath, string destinationPath);
}

internal static class ConverterFactory
{
    public static IDocumentConverter Create(OfficeFamily family) => family switch
    {
        OfficeFamily.Word => new WordConverter(),
        OfficeFamily.PowerPoint => new PowerPointConverter(),
        _ => throw new OfficeAutomationException($"No hay conversor para '{family}'."),
    };
}
