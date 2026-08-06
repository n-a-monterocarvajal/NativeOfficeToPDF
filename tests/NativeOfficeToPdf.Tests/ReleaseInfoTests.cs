using NativeOfficeToPdf.Update;

namespace NativeOfficeToPdf.Tests;

public class ReleaseInfoTests
{
    [Fact]
    public void LeeEtiquetaYUrlDeLaRespuestaDeGitHub()
    {
        const string json = """
            {
              "tag_name": "v1.4.0",
              "name": "v1.4.0 — soporte de Excel",
              "draft": false,
              "prerelease": false,
              "html_url": "https://github.com/n-a-monterocarvajal/NativeOfficeToPDF/releases/tag/v1.4.0"
            }
            """;

        Assert.True(ReleaseInfo.TryParse(json, out ReleaseInfo release));
        Assert.Equal("1.4.0", release.Version.Text);
        Assert.EndsWith("/releases/tag/v1.4.0", release.Url, StringComparison.Ordinal);
    }

    [Fact]
    public void SinHtmlUrl_caeEnLaPaginaDeReleases()
    {
        Assert.True(ReleaseInfo.TryParse("""{ "tag_name": "v2.0.0" }""", out ReleaseInfo release));
        Assert.Equal("2.0.0", release.Version.Text);
        Assert.Contains("/releases/latest", release.Url, StringComparison.Ordinal);
    }

    [Fact]
    public void BorradorNoCuenta()
    {
        Assert.False(ReleaseInfo.TryParse("""{ "tag_name": "v9.9.9", "draft": true }""", out _));
    }

    [Theory]
    [InlineData("""{ "message": "Not Found" }""")]
    [InlineData("""{ "tag_name": "sin-version" }""")]
    [InlineData("[]")]
    [InlineData("no es json")]
    [InlineData("")]
    public void RespuestasQueNoSirven_noRompen(string json)
    {
        Assert.False(ReleaseInfo.TryParse(json, out _));
    }
}
