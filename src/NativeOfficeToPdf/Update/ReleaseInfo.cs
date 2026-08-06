using System.Text.Json;

namespace NativeOfficeToPdf.Update;

/// <param name="Version">Versión publicada, leída de <c>tag_name</c>.</param>
/// <param name="Url">Página del Release en GitHub, para que el usuario decida qué hacer.</param>
internal sealed record ReleaseInfo(SemanticVersion Version, string Url)
{
    private const string FallbackUrl =
        $"https://github.com/{AppInfo.RepositoryOwner}/{AppInfo.RepositoryName}/releases/latest";

    /// <summary>
    /// Lee la respuesta de <c>GET /repos/{owner}/{repo}/releases/latest</c>. Está separado del acceso
    /// a la red para poder probarlo sin salir a internet.
    /// </summary>
    public static bool TryParse(string json, out ReleaseInfo release)
    {
        release = null!;

        try
        {
            using JsonDocument document = JsonDocument.Parse(json);
            JsonElement root = document.RootElement;

            if (root.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            if (!root.TryGetProperty("tag_name", out JsonElement tag) || tag.ValueKind != JsonValueKind.String)
            {
                return false;
            }

            if (!SemanticVersion.TryParse(tag.GetString(), out SemanticVersion version))
            {
                return false;
            }

            // Un borrador no es una versión publicada: no se avisa de él.
            if (root.TryGetProperty("draft", out JsonElement draft) &&
                draft.ValueKind == JsonValueKind.True)
            {
                return false;
            }

            string url = FallbackUrl;
            if (root.TryGetProperty("html_url", out JsonElement htmlUrl) &&
                htmlUrl.ValueKind == JsonValueKind.String &&
                !string.IsNullOrWhiteSpace(htmlUrl.GetString()))
            {
                url = htmlUrl.GetString()!;
            }

            release = new ReleaseInfo(version, url);
            return true;
        }
        catch (JsonException)
        {
            return false;
        }
    }
}
