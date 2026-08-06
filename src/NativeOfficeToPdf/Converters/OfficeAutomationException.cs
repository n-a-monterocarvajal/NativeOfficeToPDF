namespace NativeOfficeToPdf.Converters;

/// <summary>
/// Fallo atribuible a la automatización de Office: la aplicación no está instalada o registrada, se
/// negó a abrir el documento, o no produjo el PDF. Se traduce al código de salida 2048.
/// </summary>
internal sealed class OfficeAutomationException : Exception
{
    public OfficeAutomationException(string message)
        : base(message)
    {
    }

    public OfficeAutomationException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
