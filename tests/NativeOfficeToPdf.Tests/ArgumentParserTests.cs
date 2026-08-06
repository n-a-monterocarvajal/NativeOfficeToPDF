using NativeOfficeToPdf.Cli;

namespace NativeOfficeToPdf.Tests;

// Los tipos del CLI son internal (los ve la suite por InternalsVisibleTo), así que las firmas públicas
// de los métodos de prueba no pueden mencionarlos: los verbos se comparan por nombre.
public class ArgumentParserTests
{
    [Fact]
    public void SinArgumentos_esInvalido()
    {
        ParsedCommand command = ArgumentParser.Parse([]);

        Assert.False(command.IsValid);
    }

    [Fact]
    public void UnSoloArchivo_seInterpretaComoConversion()
    {
        ParsedCommand command = ArgumentParser.Parse([@"C:\docs\informe.docx"]);

        Assert.True(command.IsValid);
        Assert.Equal("Convert", command.Verb.ToString());
        Assert.Equal(@"C:\docs\informe.docx", command.Source);
        Assert.Null(command.Destination);
    }

    [Fact]
    public void OrigenYDestino_seAsignanEnOrden()
    {
        ParsedCommand command = ArgumentParser.Parse([@"C:\a.docx", @"D:\salida\a.pdf"]);

        Assert.Equal(@"C:\a.docx", command.Source);
        Assert.Equal(@"D:\salida\a.pdf", command.Destination);
    }

    [Theory]
    [InlineData("convert")]
    [InlineData("convertir")]
    public void VerboDeConversionExplicito_aceptaOrigen(string token)
    {
        ParsedCommand command = ArgumentParser.Parse([token, "a.docx"]);

        Assert.True(command.IsValid);
        Assert.Equal("Convert", command.Verb.ToString());
        Assert.Equal("a.docx", command.Source);
    }

    [Theory]
    [InlineData("install", "Install")]
    [InlineData("instalar", "Install")]
    [InlineData("uninstall", "Uninstall")]
    [InlineData("desinstalar", "Uninstall")]
    [InlineData("check-updates", "CheckUpdates")]
    [InlineData("--version", "Version")]
    [InlineData("-v", "Version")]
    [InlineData("--help", "Help")]
    [InlineData("/?", "Help")]
    public void VerbosReconocidos(string token, string expected)
    {
        ParsedCommand command = ArgumentParser.Parse([token]);

        Assert.True(command.IsValid);
        Assert.Equal(expected, command.Verb.ToString());
    }

    [Fact]
    public void VerboQueNoAceptaArgumentos_esInvalidoConUnArchivo()
    {
        ParsedCommand command = ArgumentParser.Parse(["install", "a.docx"]);

        Assert.False(command.IsValid);
    }

    [Fact]
    public void Banderas_seReconocenEnCualquierPosicion()
    {
        ParsedCommand command = ArgumentParser.Parse(["--overwrite", "a.docx", "--open", "-q"]);

        Assert.True(command.IsValid);
        Assert.True(command.Overwrite);
        Assert.True(command.Open);
        Assert.True(command.Quiet);
        Assert.Equal("a.docx", command.Source);
    }

    [Fact]
    public void OpcionDesconocida_esInvalida()
    {
        ParsedCommand command = ArgumentParser.Parse(["a.docx", "--comprimir"]);

        Assert.False(command.IsValid);
    }

    [Fact]
    public void TercerArgumentoPosicional_esInvalido()
    {
        ParsedCommand command = ArgumentParser.Parse(["a.docx", "b.pdf", "c.pdf"]);

        Assert.False(command.IsValid);
    }

    [Fact]
    public void ArchivoQueSeLlamaComoUnVerbo_ganaElArchivo()
    {
        // Un archivo llamado "install" no debe disparar la instalación: el uso principal de la
        // herramienta es recibir una ruta como primer argumento.
        const string name = "install";
        File.WriteAllText(name, string.Empty);

        try
        {
            ParsedCommand command = ArgumentParser.Parse([name]);

            Assert.True(command.IsValid);
            Assert.Equal("Convert", command.Verb.ToString());
            Assert.Equal(name, command.Source);
        }
        finally
        {
            File.Delete(name);
        }
    }
}
