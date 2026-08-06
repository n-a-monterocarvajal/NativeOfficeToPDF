using NativeOfficeToPdf.Cli;

namespace NativeOfficeToPdf.Tests;

public class ConversionPathsTests
{
    [Fact]
    public void SinDestino_dejaElPdfJuntoAlOriginal()
    {
        string result = ConversionPaths.ResolveDestination(@"C:\docs\Informe final.docx", null);

        Assert.Equal(@"C:\docs\Informe final.pdf", result);
    }

    [Fact]
    public void SinDestino_conservaOtrosPuntosDelNombre()
    {
        string result = ConversionPaths.ResolveDestination(@"C:\docs\acta.2026.03.pptx", null);

        Assert.Equal(@"C:\docs\acta.2026.03.pdf", result);
    }

    [Fact]
    public void DestinoConExtension_seRespeta()
    {
        string result = ConversionPaths.ResolveDestination(@"C:\docs\a.docx", @"D:\salida\otro.pdf");

        Assert.Equal(@"D:\salida\otro.pdf", result);
    }

    [Fact]
    public void DestinoSinExtension_recibePdf()
    {
        string result = ConversionPaths.ResolveDestination(@"C:\docs\a.docx", @"D:\salida\otro");

        Assert.Equal(@"D:\salida\otro.pdf", result);
    }

    [Fact]
    public void DestinoTerminadoEnSeparador_seTrataComoCarpeta()
    {
        string result = ConversionPaths.ResolveDestination(@"C:\docs\a.docx", @"D:\salida\");

        Assert.Equal(@"D:\salida\a.pdf", result);
    }

    [Fact]
    public void DestinoQueEsUnaCarpetaExistente_seTrataComoCarpeta()
    {
        string folder = Directory.CreateTempSubdirectory("nativeofficetopdf").FullName;

        try
        {
            string result = ConversionPaths.ResolveDestination(@"C:\docs\a.docx", folder);

            Assert.Equal(Path.Combine(folder, "a.pdf"), result);
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void OrigenRelativo_seResuelveAAbsoluto()
    {
        string result = ConversionPaths.ResolveDestination("informe.docx", null);

        Assert.True(Path.IsPathRooted(result));
        Assert.Equal("informe.pdf", Path.GetFileName(result));
    }
}
