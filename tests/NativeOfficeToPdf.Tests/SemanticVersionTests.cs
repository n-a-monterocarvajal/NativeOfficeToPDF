using NativeOfficeToPdf.Update;

namespace NativeOfficeToPdf.Tests;

public class SemanticVersionTests
{
    [Theory]
    [InlineData("1.2.3", "1.2.3")]
    [InlineData("v1.2.3", "1.2.3")]
    [InlineData("V1.2.3", "1.2.3")]
    [InlineData("1.2", "1.2.0")]
    [InlineData("1.2.3.4", "1.2.3")]
    [InlineData("1.2.3+abc123", "1.2.3")]
    [InlineData("0.7.0-alpha", "0.7.0-alpha")]
    [InlineData("  1.0.0  ", "1.0.0")]
    public void Normaliza(string input, string expected)
    {
        Assert.True(SemanticVersion.TryParse(input, out SemanticVersion version));
        Assert.Equal(expected, version.Text);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("no-es-version")]
    [InlineData("1.x.3")]
    [InlineData("-1.0.0")]
    [InlineData("1.0.0-")]
    public void RechazaLoQueNoEsVersion(string? input)
    {
        Assert.False(SemanticVersion.TryParse(input, out _));
    }

    [Theory]
    [InlineData("1.0.1", "1.0.0")]
    [InlineData("1.1.0", "1.0.9")]
    [InlineData("2.0.0", "1.99.99")]
    [InlineData("1.0.0", "1.0.0-alpha")]
    [InlineData("1.0.0-beta", "1.0.0-alpha")]
    [InlineData("1.0.0-alpha.2", "1.0.0-alpha.1")]
    [InlineData("1.0.0-alpha.beta", "1.0.0-alpha.11")]
    [InlineData("1.0.0-alpha.1", "1.0.0-alpha")]
    public void ComparaPrecedencia(string mayor, string menor)
    {
        Assert.True(SemanticVersion.TryParse(mayor, out SemanticVersion left));
        Assert.True(SemanticVersion.TryParse(menor, out SemanticVersion right));

        Assert.True(left.CompareTo(right) > 0);
        Assert.True(right.CompareTo(left) < 0);
    }

    [Fact]
    public void VersionesEquivalentes_sonIguales()
    {
        Assert.True(SemanticVersion.TryParse("v1.2.3+meta", out SemanticVersion left));
        Assert.True(SemanticVersion.TryParse("1.2.3", out SemanticVersion right));

        Assert.Equal(left, right);
        Assert.Equal(0, left.CompareTo(right));
    }
}
