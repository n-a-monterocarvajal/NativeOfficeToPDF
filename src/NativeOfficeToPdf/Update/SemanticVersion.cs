using System.Globalization;

namespace NativeOfficeToPdf.Update;

/// <summary>
/// Versión semántica reducida a lo que este proyecto necesita: comparar la etiqueta del último Release
/// con la versión compilada. Acepta la <c>v</c> inicial de las etiquetas de Git e ignora los metadatos
/// de compilación (<c>+sha</c>), que por especificación no participan en la precedencia.
/// </summary>
internal sealed class SemanticVersion : IComparable<SemanticVersion>, IEquatable<SemanticVersion>
{
    private SemanticVersion(int major, int minor, int patch, string[] preRelease, string text)
    {
        Major = major;
        Minor = minor;
        Patch = patch;
        PreRelease = preRelease;
        Text = text;
    }

    public int Major { get; }

    public int Minor { get; }

    public int Patch { get; }

    /// <summary>Identificadores de prelanzamiento, vacío para una versión estable.</summary>
    public IReadOnlyList<string> PreRelease { get; }

    /// <summary>Texto normalizado, sin la <c>v</c> inicial ni los metadatos de compilación.</summary>
    public string Text { get; }

    public static bool TryParse(string? value, out SemanticVersion version)
    {
        version = null!;

        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        string text = value.Trim();
        if (text.StartsWith('v') || text.StartsWith('V'))
        {
            text = text[1..];
        }

        int plus = text.IndexOf('+');
        if (plus >= 0)
        {
            text = text[..plus];
        }

        string[] preRelease = [];
        int dash = text.IndexOf('-');
        if (dash >= 0)
        {
            string tail = text[(dash + 1)..];
            text = text[..dash];
            if (tail.Length == 0)
            {
                return false;
            }

            preRelease = tail.Split('.');
        }

        string[] parts = text.Split('.');
        if (parts.Length is < 1 or > 4)
        {
            return false;
        }

        // Se aceptan hasta cuatro componentes porque AssemblyVersion es X.Y.Z.W; el cuarto se ignora.
        int[] numbers = new int[3];
        for (int i = 0; i < 3; i++)
        {
            if (i >= parts.Length)
            {
                numbers[i] = 0;
                continue;
            }

            if (!int.TryParse(parts[i], NumberStyles.None, CultureInfo.InvariantCulture, out int parsed))
            {
                return false;
            }

            numbers[i] = parsed;
        }

        string normalized = $"{numbers[0]}.{numbers[1]}.{numbers[2]}";
        if (preRelease.Length > 0)
        {
            normalized += "-" + string.Join('.', preRelease);
        }

        version = new SemanticVersion(numbers[0], numbers[1], numbers[2], preRelease, normalized);
        return true;
    }

    public int CompareTo(SemanticVersion? other)
    {
        if (other is null)
        {
            return 1;
        }

        int result = Major.CompareTo(other.Major);
        if (result != 0)
        {
            return result;
        }

        result = Minor.CompareTo(other.Minor);
        if (result != 0)
        {
            return result;
        }

        result = Patch.CompareTo(other.Patch);
        if (result != 0)
        {
            return result;
        }

        // Una versión con prelanzamiento precede siempre a la estable equivalente: 1.0.0-alpha < 1.0.0.
        if (PreRelease.Count == 0 && other.PreRelease.Count == 0)
        {
            return 0;
        }

        if (PreRelease.Count == 0)
        {
            return 1;
        }

        if (other.PreRelease.Count == 0)
        {
            return -1;
        }

        int shared = Math.Min(PreRelease.Count, other.PreRelease.Count);
        for (int i = 0; i < shared; i++)
        {
            result = CompareIdentifiers(PreRelease[i], other.PreRelease[i]);
            if (result != 0)
            {
                return result;
            }
        }

        return PreRelease.Count.CompareTo(other.PreRelease.Count);
    }

    private static int CompareIdentifiers(string left, string right)
    {
        bool leftNumeric = int.TryParse(left, NumberStyles.None, CultureInfo.InvariantCulture, out int leftValue);
        bool rightNumeric = int.TryParse(right, NumberStyles.None, CultureInfo.InvariantCulture, out int rightValue);

        if (leftNumeric && rightNumeric)
        {
            return leftValue.CompareTo(rightValue);
        }

        // Los identificadores numéricos tienen menor precedencia que los alfanuméricos.
        if (leftNumeric)
        {
            return -1;
        }

        if (rightNumeric)
        {
            return 1;
        }

        return string.CompareOrdinal(left, right);
    }

    public bool Equals(SemanticVersion? other) => CompareTo(other) == 0;

    public override bool Equals(object? obj) => obj is SemanticVersion other && Equals(other);

    public override int GetHashCode() => Text.GetHashCode(StringComparison.Ordinal);

    public override string ToString() => Text;
}
