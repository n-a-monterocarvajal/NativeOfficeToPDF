using System.Diagnostics;
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
    public static string ExecutablePath { get; } = ReadExecutablePath();

    /// <summary>Carpeta de estado por usuario. No requiere permisos especiales.</summary>
    public static string StateDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Name);

    private static string ReadVersion()
    {
        Assembly assembly = typeof(AppInfo).Assembly;

        string? informational = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;

        if (!string.IsNullOrWhiteSpace(informational))
        {
            // El SDK agrega "+<sha>" cuando el repositorio tiene metadatos de origen.
            int plus = informational.IndexOf('+');
            return plus >= 0 ? informational[..plus] : informational;
        }

        return assembly.GetName().Version?.ToString(3) ?? "0.0.0";
    }

    private static string ReadExecutablePath()
    {
        string? path = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(path))
        {
            return path;
        }

        using Process current = Process.GetCurrentProcess();
        return current.MainModule?.FileName ?? Path.Combine(AppContext.BaseDirectory, Name + ".exe");
    }
}
