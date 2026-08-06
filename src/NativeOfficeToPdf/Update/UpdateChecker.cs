using System.Net.Http;
using System.Net.Http.Headers;

namespace NativeOfficeToPdf.Update;

/// <summary>
/// Chequeo de actualizaciones contra los Releases del repositorio.
/// <para>
/// Tres reglas lo gobiernan: <b>nunca bloquea</b> (cinco segundos de tope y cualquier error se traga),
/// <b>nunca molesta dos veces</b> (un chequeo cada veinticuatro horas y un solo aviso por versión), y
/// <b>nunca instala nada</b> (solo ofrece abrir la página del Release). Si el repositorio es privado,
/// la API responde 404 sin credenciales y el chequeo simplemente no encuentra nada — el resto de la
/// herramienta funciona igual.
/// </para>
/// </summary>
internal static class UpdateChecker
{
    private static readonly TimeSpan CheckInterval = TimeSpan.FromHours(24);
    private static readonly TimeSpan NetworkTimeout = TimeSpan.FromSeconds(5);

    private static readonly Uri LatestReleaseUri = new(
        $"https://api.github.com/repos/{AppInfo.RepositoryOwner}/{AppInfo.RepositoryName}/releases/latest");

    /// <summary>Resultado de un chequeo.</summary>
    /// <param name="Release">Versión publicada más reciente, o nulo si no se pudo consultar.</param>
    /// <param name="IsNewer">Verdadero si esa versión es posterior a la que está corriendo.</param>
    /// <param name="AlreadyNotified">Verdadero si ya se avisó de esa versión antes.</param>
    internal sealed record Result(ReleaseInfo? Release, bool IsNewer, bool AlreadyNotified)
    {
        public static Result None { get; } = new(null, false, false);
    }

    /// <param name="force">
    /// Verdadero para saltarse la ventana de veinticuatro horas. Lo usa el verbo <c>check-updates</c>,
    /// donde el usuario pidió el chequeo explícitamente.
    /// </param>
    public static Result Check(bool force)
    {
        try
        {
            UpdateStateData state = UpdateState.Load();

            if (!force && state.LastCheckUtc is { } last &&
                DateTimeOffset.UtcNow - last < CheckInterval)
            {
                return Result.None;
            }

            string? json = Download();
            state.LastCheckUtc = DateTimeOffset.UtcNow;
            UpdateState.Save(state);

            if (json is null || !ReleaseInfo.TryParse(json, out ReleaseInfo release))
            {
                return Result.None;
            }

            if (!SemanticVersion.TryParse(AppInfo.Version, out SemanticVersion current))
            {
                return new Result(release, IsNewer: false, AlreadyNotified: false);
            }

            bool isNewer = release.Version.CompareTo(current) > 0;
            bool alreadyNotified = string.Equals(
                state.LastNotifiedVersion, release.Version.Text, StringComparison.Ordinal);

            return new Result(release, isNewer, alreadyNotified);
        }
        catch (Exception)
        {
            // Deliberadamente amplio: este chequeo es accesorio y jamás debe alterar el resultado ni el
            // código de salida de una conversión.
            return Result.None;
        }
    }

    /// <summary>Recuerda que ya se avisó de esta versión, para no repetir el aviso.</summary>
    public static void MarkNotified(ReleaseInfo release)
    {
        UpdateStateData state = UpdateState.Load();
        state.LastNotifiedVersion = release.Version.Text;
        UpdateState.Save(state);
    }

    private static string? Download()
    {
        using var client = new HttpClient { Timeout = NetworkTimeout };
        using var request = new HttpRequestMessage(HttpMethod.Get, LatestReleaseUri);

        // GitHub exige User-Agent en todas las peticiones a su API.
        request.Headers.UserAgent.Add(new ProductInfoHeaderValue(AppInfo.Name, AppInfo.Version));
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");

        using HttpResponseMessage response = client.Send(request);
        return response.IsSuccessStatusCode
            ? response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            : null;
    }
}
