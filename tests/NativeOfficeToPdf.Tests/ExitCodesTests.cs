using System.Reflection;
using System.Text.RegularExpressions;
using NativeOfficeToPdf.Cli;

namespace NativeOfficeToPdf.Tests;

public class ExitCodesTests
{
    private static int[] Declarados() =>
        [.. typeof(ExitCodes)
            .GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
            .Where(campo => campo is { IsLiteral: true, IsInitOnly: false })
            .Select(campo => (int)campo.GetRawConstantValue()!)];

    [Fact]
    public void SonPropiosYCorrelativos()
    {
        // Numeración propia: correlativa desde 0, sin huecos. El proyecto no hereda el esquema de
        // ninguna otra herramienta de conversión.
        Assert.Equal(Enumerable.Range(0, Declarados().Length), Declarados().Order());
    }

    /// <summary>
    /// Los códigos son contrato público del CLI, y el README es donde los lee quien escribe scripts.
    /// Esta prueba existe para que agregar un código y olvidar la tabla rompa el CI.
    /// </summary>
    [Fact]
    public void LaTablaDelReadmeListaExactamenteLosCodigosDeclarados()
    {
        string readme = File.ReadAllText(Path.Combine(RepositoryLayout.Root, "README.md"));

        int[] documentados =
            [.. Regex.Matches(readme, @"^\|\s*(?<codigo>\d+)\s*\|", RegexOptions.Multiline)
                .Select(coincidencia => int.Parse(coincidencia.Groups["codigo"].Value))
                .Distinct()
                .Order()];

        Assert.Equal(Declarados().Order(), documentados);
    }
}
