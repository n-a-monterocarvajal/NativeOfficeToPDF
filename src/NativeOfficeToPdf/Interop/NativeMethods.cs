using System.Runtime.InteropServices;

namespace NativeOfficeToPdf.Interop;

/// <summary>
/// P/Invoke mínimo. Se usa DllImport clásico en vez de LibraryImport para no tener que habilitar
/// bloques unsafe en el proyecto: son cinco funciones y ninguna está en un camino caliente.
/// </summary>
internal static class NativeMethods
{
    /// <summary>Consola del proceso padre, para <see cref="AttachConsole"/>.</summary>
    internal const uint AttachParentProcess = 0xFFFFFFFF;

    internal const uint MbOk = 0x00000000;
    internal const uint MbYesNo = 0x00000004;
    internal const uint MbIconError = 0x00000010;
    internal const uint MbIconQuestion = 0x00000020;
    internal const uint MbIconInformation = 0x00000040;
    internal const uint MbSetForeground = 0x00010000;

    internal const int IdYes = 6;

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool AttachConsole(uint processId);

    [DllImport("kernel32.dll")]
    internal static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = false)]
    internal static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

    internal const int TdcbfOk = 0x0001;
    internal const int TdcbfYes = 0x0002;
    internal const int TdcbfNo = 0x0004;

    internal const int IdNo = 7;

    /// <summary>
    /// Los iconos de <c>TaskDialog</c> se pasan como identificadores de recurso negativos disfrazados
    /// de puntero (<c>MAKEINTRESOURCEW</c>), no como cadenas.
    /// </summary>
    internal static readonly IntPtr TdWarningIcon = new(0xFFFF);

    internal static readonly IntPtr TdErrorIcon = new(0xFFFE);

    internal static readonly IntPtr TdInformationIcon = new(0xFFFD);

    /// <summary>
    /// El diálogo moderno de Windows: título, instrucción principal en grande y cuerpo. Solo existe en
    /// la versión 6 de comctl32, que el proceso obtiene por el manifiesto (<c>app.manifest</c>); sin
    /// él, esta entrada de función no se resuelve. <see cref="Cli.UserOutput"/> cae a
    /// <see cref="MessageBoxW"/> si eso ocurre.
    /// </summary>
    [DllImport("comctl32.dll", CharSet = CharSet.Unicode, SetLastError = false, PreserveSig = true)]
    internal static extern int TaskDialog(
        IntPtr hwndParent,
        IntPtr hInstance,
        string windowTitle,
        string mainInstruction,
        string? content,
        int commonButtons,
        IntPtr icon,
        out int pressedButton);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
    internal static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    /// <summary>
    /// Recupera del Running Object Table una instancia ya en marcha del servidor COM. Es el reemplazo
    /// de <c>Marshal.GetActiveObject</c>, que solo existe en .NET Framework.
    /// </summary>
    [DllImport("oleaut32.dll", ExactSpelling = true, PreserveSig = true)]
    internal static extern int GetActiveObject(
        ref Guid clsid,
        IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object instance);
}
