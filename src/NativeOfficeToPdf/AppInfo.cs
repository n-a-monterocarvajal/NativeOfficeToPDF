using System.Reflection;

namespace NativeOfficeToPdf;

internal static class AppInfo
{
    public const string Name = "NativeOfficeToPdf";

    /// <summary>Repositorio del que se leen los Releases para el chequeo de actualizaciones.</summary>
    public const string RepositoryOwner = "n-a-monterocarvajal";
    public const string RepositoryName = "NativeOfficeToPDF";

    /// <summary>Versión del producto, tal como quedó grabada en el ensamblado al compilar.</summary>
    public static string Version { get; } = ReadVersion();

    /// <summary>Ruta del ejecutable en marcha. Es lo que se escribe en el registro al instalar.</summary>
    public static string ExecutablePath { get; } = Environment.ProcessPath!;

    /// <summary>Carpeta de estado por usuario. No requiere permisos especiales.</summary>
    public static string StateDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Name);

    // El SDK siempre graba <Version> como InformationalVersion, y le agrega "+<sha>" cuando el
    // repositorio tiene metadatos de origen.
    private static string ReadVersion() =>
        typeof(AppInfo).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion
            .Split('+')[0];
}
