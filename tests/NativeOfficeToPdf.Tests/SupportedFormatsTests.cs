using System.Text.RegularExpressions;
using NativeOfficeToPdf.Converters;
using NativeOfficeToPdf.Shell;

namespace NativeOfficeToPdf.Tests;

public class SupportedFormatsTests
{
    [Theory]
    [InlineData(@"C:\a.docx", "Word")]
    [InlineData(@"C:\a.DOCX", "Word")]
    [InlineData(@"C:\a.doc", "Word")]
    [InlineData(@"C:\a.docm", "Word")]
    [InlineData(@"C:\a.pptx", "PowerPoint")]
    [InlineData(@"C:\a.ppt", "PowerPoint")]
    [InlineData(@"C:\a.pptm", "PowerPoint")]
    public void ReconoceLosFormatosSoportados(string path, string expectedFamily)
    {
        Assert.True(SupportedFormats.TryGetFamily(path, out OfficeFamily family));
        Assert.Equal(expectedFamily, family.ToString());
    }

    [Theory]
    [InlineData(@"C:\a.pdf")]
    [InlineData(@"C:\a.xlsx")]
    [InlineData(@"C:\a.txt")]
    [InlineData(@"C:\sin-extension")]
    public void RechazaLoQueNoSabeConvertir(string path)
    {
        Assert.False(SupportedFormats.IsSupported(path));
    }

    [Fact]
    public void LaRutaDelRegistroApuntaALaRamaDeUsuario()
    {
        string path = ContextMenuRegistrar.KeyPathFor(".docx");

        Assert.Equal(@"Software\Classes\SystemFileAssociations\.docx\shell\ConvertToPdf", path);
    }

    /// <summary>
    /// El instalador de Inno Setup repite la lista de extensiones en su propia sintaxis. Esta prueba
    /// existe para que agregar un formato en <c>SupportedFormats</c> y olvidarlo en el instalador
    /// rompa el CI, en vez de producir una entrada de menú que solo aparece a medias.
    /// </summary>
    [Fact]
    public void ElInstaladorRegistraExactamenteLasMismasExtensiones()
    {
        string script = File.ReadAllText(RepositoryLayout.InstallerScript);
        Match match = Regex.Match(
            script,
            @"^#define\s+SupportedExtensions\s+""(?<lista>[^""]*)""",
            RegexOptions.Multiline);

        Assert.True(match.Success, "El .iss debe declarar #define SupportedExtensions.");

        string[] fromInstaller = match.Groups["lista"].Value
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        Assert.Equal(
            SupportedFormats.Extensions.Order(StringComparer.Ordinal),
            fromInstaller.Order(StringComparer.Ordinal));

        // Y cada extensión declarada tiene que aparecer de verdad en la sección [Registry]: el
        // #define por sí solo podría quedar como adorno.
        foreach (string extension in fromInstaller)
        {
            Assert.Contains(
                $@"Subkey: ""Software\Classes\SystemFileAssociations\{extension}\shell\",
                script,
                StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ElInstaladorUsaElMismoNombreDeVerbo()
    {
        string script = File.ReadAllText(RepositoryLayout.InstallerScript);

        Assert.Contains($@"#define VerbKeyName ""{ContextMenuRegistrar.VerbKeyName}""", script, StringComparison.Ordinal);
        Assert.Contains($@"#define VerbDisplayName ""{ContextMenuRegistrar.VerbDisplayName}""", script, StringComparison.Ordinal);
    }
}
