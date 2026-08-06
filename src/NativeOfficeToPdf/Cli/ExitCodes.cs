namespace NativeOfficeToPdf.Cli;

/// <summary>
/// Códigos de salida. Los valores 0, 1, 8, 32, 64 y 2048 siguen la convención de OfficeToPDF, para que
/// cualquier script de verificación existente siga funcionando sin cambios.
/// </summary>
internal static class ExitCodes
{
    public const int Success = 0;
    public const int Failure = 1;

    /// <summary>Solo lo devuelve <c>check-updates</c>: hay una versión más nueva publicada.</summary>
    public const int UpdateAvailable = 4;

    public const int InvalidArguments = 8;
    public const int UnsupportedExtension = 32;
    public const int SourceNotFound = 64;
    public const int OfficeAutomationError = 2048;
}
