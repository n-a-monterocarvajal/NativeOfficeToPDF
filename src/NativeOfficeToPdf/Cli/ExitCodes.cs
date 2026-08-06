namespace NativeOfficeToPdf.Cli;

/// <summary>
/// Códigos de salida del programa. Numeración propia y correlativa: este proyecto no hereda el
/// esquema de ninguna otra herramienta de conversión.
/// <para>
/// Son parte del contrato público del CLI: cambiar un valor rompe los scripts de quien lo use, así
/// que un cambio acá va acompañado de la tabla del README, el texto de ayuda y el paso de humo del CI.
/// </para>
/// </summary>
internal static class ExitCodes
{
    public const int Success = 0;
    public const int Failure = 1;
    public const int InvalidArguments = 2;
    public const int UnsupportedExtension = 3;
    public const int SourceNotFound = 4;
    public const int OfficeAutomationError = 5;

    /// <summary>Solo lo devuelve <c>check-updates</c>: hay una versión más nueva publicada.</summary>
    public const int UpdateAvailable = 6;
}
