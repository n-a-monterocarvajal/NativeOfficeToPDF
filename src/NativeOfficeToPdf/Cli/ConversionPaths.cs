namespace NativeOfficeToPdf.Cli;

internal static class ConversionPaths
{
    /// <summary>
    /// Resuelve la ruta del PDF de salida. Office resuelve las rutas relativas contra su propio
    /// directorio de trabajo, así que acá siempre se devuelve una ruta absoluta.
    /// </summary>
    /// <param name="sourcePath">Ruta del documento de origen, relativa o absoluta.</param>
    /// <param name="destination">
    /// Destino pedido por el usuario. Puede ser un archivo, un directorio existente, o nulo para
    /// dejar el PDF junto al original con el mismo nombre.
    /// </param>
    public static string ResolveDestination(string sourcePath, string? destination)
    {
        string source = Path.GetFullPath(sourcePath);
        string defaultName = Path.GetFileNameWithoutExtension(source) + ".pdf";

        if (string.IsNullOrWhiteSpace(destination))
        {
            return Path.Combine(Path.GetDirectoryName(source) ?? string.Empty, defaultName);
        }

        string requested = destination.Trim();
        bool looksLikeDirectory =
            requested.EndsWith(Path.DirectorySeparatorChar) ||
            requested.EndsWith(Path.AltDirectorySeparatorChar) ||
            Directory.Exists(requested);

        if (looksLikeDirectory)
        {
            return Path.Combine(Path.GetFullPath(requested), defaultName);
        }

        string full = Path.GetFullPath(requested);
        return string.IsNullOrEmpty(Path.GetExtension(full)) ? full + ".pdf" : full;
    }
}
