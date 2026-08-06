using Microsoft.Win32;
using NativeOfficeToPdf.Converters;

namespace NativeOfficeToPdf.Shell;

/// <summary>
/// Registro del verbo "Convertir a PDF" en el menú contextual de Explorer.
/// <para>
/// Todo se escribe bajo <c>HKEY_CURRENT_USER</c>. Windows fusiona <c>HKCU\Software\Classes</c> con
/// <c>HKLM\Software\Classes</c> para resolver el menú contextual, y la rama de usuario tiene prioridad
/// — así que se consigue exactamente el mismo resultado que escribiendo por máquina, pero sin UAC ni
/// en la instalación ni en el uso. La contrapartida: en un equipo compartido hay que instalar con cada
/// cuenta.
/// </para>
/// <para>
/// El instalador de Inno Setup escribe estas mismas claves por su cuenta (ver
/// <c>installer/NativeOfficeToPdf.iss</c>); esta clase es la vía para uso portable o despliegue
/// scriptado, y es la que cubren las pruebas.
/// </para>
/// </summary>
internal static class ContextMenuRegistrar
{
    /// <summary>Nombre de la clave del verbo. Debe coincidir con el del instalador.</summary>
    public const string VerbKeyName = "ConvertToPdf";

    /// <summary>Texto que ve el usuario en el menú contextual. Debe coincidir con el del instalador.</summary>
    public const string VerbDisplayName = "Convertir a PDF";

    private const string ClassesRoot = @"Software\Classes\SystemFileAssociations";

    public static string KeyPathFor(string extension) =>
        $@"{ClassesRoot}\{extension}\shell\{VerbKeyName}";

    /// <summary>
    /// Registra el verbo apuntando a <paramref name="executablePath"/>. No copia nada: registra el
    /// ejecutable donde esté, que es lo que se quiere en uso portable.
    /// </summary>
    public static void Install(string executablePath)
    {
        string command = $"\"{executablePath}\" \"%1\"";

        foreach (string extension in SupportedFormats.Extensions)
        {
            using RegistryKey verb = Registry.CurrentUser.CreateSubKey(KeyPathFor(extension), writable: true);
            verb.SetValue("MUIVerb", VerbDisplayName, RegistryValueKind.String);
            verb.SetValue("Icon", executablePath, RegistryValueKind.String);

            using RegistryKey commandKey = verb.CreateSubKey("command", writable: true);
            commandKey.SetValue(null, command, RegistryValueKind.String);
        }
    }

    /// <summary>Borra el verbo de las seis extensiones. Es idempotente.</summary>
    public static void Uninstall()
    {
        foreach (string extension in SupportedFormats.Extensions)
        {
            Registry.CurrentUser.DeleteSubKeyTree(KeyPathFor(extension), throwOnMissingSubKey: false);
        }
    }

    /// <summary>Verdadero si el verbo está registrado para todas las extensiones soportadas.</summary>
    public static bool IsInstalled()
    {
        foreach (string extension in SupportedFormats.Extensions)
        {
            using RegistryKey? verb = Registry.CurrentUser.OpenSubKey(KeyPathFor(extension));
            if (verb is null)
            {
                return false;
            }
        }

        return true;
    }
}
