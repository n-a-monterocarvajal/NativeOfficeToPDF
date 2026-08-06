namespace NativeOfficeToPdf.Tests;

/// <summary>
/// Localiza la raíz del repositorio a partir del directorio de salida de las pruebas, para poder leer
/// archivos que no son código compilado (el script del instalador, por ejemplo).
/// </summary>
internal static class RepositoryLayout
{
    public static string Root { get; } = FindRoot();

    public static string InstallerScript => Path.Combine(Root, "installer", "NativeOfficeToPdf.iss");

    private static string FindRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);

        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "NativeOfficeToPdf.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException(
            $"No se encontró la raíz del repositorio desde {AppContext.BaseDirectory}.");
    }
}
