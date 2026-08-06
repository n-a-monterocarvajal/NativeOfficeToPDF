using System.Text.Json;
using System.Text.Json.Serialization;

namespace NativeOfficeToPdf.Update;

internal sealed class UpdateStateData
{
    public DateTimeOffset? LastCheckUtc { get; set; }

    /// <summary>Última versión de la que ya se avisó, para no repetir el aviso en cada conversión.</summary>
    public string? LastNotifiedVersion { get; set; }
}

[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(UpdateStateData))]
internal sealed partial class UpdateStateContext : JsonSerializerContext
{
}

/// <summary>
/// Estado del chequeo de actualizaciones, en <c>%LOCALAPPDATA%\NativeOfficeToPdf\update-check.json</c>.
/// Todo el acceso es tolerante a fallos: si el archivo no existe, está corrupto o el disco lo rechaza,
/// se sigue con el estado por omisión. Nada de esto debe poder romper una conversión.
/// </summary>
internal static class UpdateState
{
    public static string FilePath => Path.Combine(AppInfo.StateDirectory, "update-check.json");

    public static UpdateStateData Load()
    {
        try
        {
            if (!File.Exists(FilePath))
            {
                return new UpdateStateData();
            }

            string json = File.ReadAllText(FilePath);
            return JsonSerializer.Deserialize(json, UpdateStateContext.Default.UpdateStateData)
                ?? new UpdateStateData();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            return new UpdateStateData();
        }
    }

    public static void Save(UpdateStateData state)
    {
        try
        {
            Directory.CreateDirectory(AppInfo.StateDirectory);
            string json = JsonSerializer.Serialize(state, UpdateStateContext.Default.UpdateStateData);
            File.WriteAllText(FilePath, json);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // El chequeo de actualizaciones es accesorio: si no se puede recordar, se vuelve a chequear
            // la próxima vez y ya.
        }
    }
}
