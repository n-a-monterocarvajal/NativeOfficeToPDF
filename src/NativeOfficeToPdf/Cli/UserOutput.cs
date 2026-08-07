using NativeOfficeToPdf.Interop;

namespace NativeOfficeToPdf.Cli;

/// <summary>
/// El puente entre ser un CLI y ser una app sin consola.
/// <para>
/// El binario se compila como <c>WinExe</c> para que el menú contextual no haga parpadear una ventana
/// negra. Eso, por sí solo, dejaría a la herramienta muda desde la línea de comandos, así que al
/// arrancar se intenta enganchar la consola del proceso padre: si el usuario lo invocó desde
/// PowerShell o cmd, escribe ahí como cualquier CLI; si lo invocó Explorer, no hay consola y los
/// errores salen por un cuadro de diálogo. El éxito nunca muestra nada.
/// </para>
/// </summary>
internal sealed class UserOutput
{
    private const string Caption = "Convertir a PDF";

    /// <summary>
    /// Válvula de escape para uso desatendido: con <c>NATIVEOFFICETOPDF_NO_DIALOGS=1</c> la herramienta
    /// no abre ningún cuadro de diálogo aunque no tenga consola. La usan el CI y cualquier despliegue
    /// automatizado, donde un diálogo modal sería un cuelgue.
    /// </summary>
    private const string NoDialogsVariable = "NATIVEOFFICETOPDF_NO_DIALOGS";

    private readonly bool dialogsAllowed;

    public UserOutput()
    {
        HasConsole = NativeMethods.GetConsoleWindow() != IntPtr.Zero ||
                     NativeMethods.AttachConsole(NativeMethods.AttachParentProcess);

        string? suppress = Environment.GetEnvironmentVariable(NoDialogsVariable);
        dialogsAllowed = !string.Equals(suppress, "1", StringComparison.Ordinal) &&
                         !string.Equals(suppress, "true", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Verdadero si hay una consola donde escribir; falso si la invocación vino de Explorer.</summary>
    public bool HasConsole { get; }

    public void Info(string message)
    {
        if (HasConsole)
        {
            Console.Out.WriteLine(message);
        }
    }

    /// <summary>Mensaje informativo que sí debe verse aunque no haya consola.</summary>
    public void Notice(string message)
    {
        if (HasConsole || !dialogsAllowed)
        {
            Console.Out.WriteLine(message);
            return;
        }

        ShowDialog(message, NativeMethods.TdInformationIcon, NativeMethods.MbIconInformation,
                   NativeMethods.TdcbfOk, NativeMethods.MbOk);
    }

    public void Error(string message)
    {
        if (HasConsole || !dialogsAllowed)
        {
            Console.Error.WriteLine(message);
            return;
        }

        ShowDialog(message, NativeMethods.TdErrorIcon, NativeMethods.MbIconError,
                   NativeMethods.TdcbfOk, NativeMethods.MbOk);
    }

    /// <summary>
    /// Error dirigido a quien invoca desde la línea de comandos: menciona opciones del CLI, así que
    /// nunca debe acabar en un cuadro de diálogo. Quien llega por el menú contextual no sabe qué es
    /// <c>--overwrite</c> y, en los casos donde se usa, ya ha respondido él mismo a la pregunta: el
    /// código de salida es el que informa.
    /// </summary>
    public void ConsoleError(string message)
    {
        if (HasConsole)
        {
            Console.Error.WriteLine(message);
        }
    }

    /// <summary>
    /// Pregunta sí/no. Sin consola se pregunta con un cuadro de diálogo; con consola no se pregunta
    /// nada — un CLI no debe quedarse esperando en un script — y se devuelve
    /// <paramref name="consoleAnswer"/>.
    /// </summary>
    public bool Confirm(string question, bool consoleAnswer)
    {
        if (HasConsole || !dialogsAllowed)
        {
            return consoleAnswer;
        }

        // Sin icono: el diálogo moderno no tiene equivalente del signo de interrogación —Windows lo
        // retiró— y una pregunta con icono de advertencia dramatiza de más.
        int result = ShowDialog(
            question,
            IntPtr.Zero,
            NativeMethods.MbIconQuestion,
            NativeMethods.TdcbfYes | NativeMethods.TdcbfNo,
            NativeMethods.MbYesNo);

        return result == NativeMethods.IdYes;
    }

    /// <summary>
    /// Dibuja el diálogo con <c>TaskDialog</c> —el aspecto actual de Windows— y cae a
    /// <c>MessageBoxW</c> si no está disponible: comctl32 v6 llega por el manifiesto, y si algún día
    /// falta, la herramienta debe seguir preguntando en vez de romperse.
    /// <para>
    /// La primera línea del mensaje se usa como instrucción principal (el texto grande) y el resto
    /// como cuerpo. Así los mensajes se escriben en el sitio de la llamada como un texto normal y no
    /// hay que partir cada uno en dos parámetros.
    /// </para>
    /// </summary>
    private static int ShowDialog(string message, IntPtr taskDialogIcon, uint messageBoxIcon,
                                  int taskDialogButtons, uint messageBoxButtons)
    {
        int corte = message.IndexOf('\n');
        string encabezado = corte < 0 ? message : message[..corte].TrimEnd('\r');
        string? cuerpo = corte < 0 ? null : message[(corte + 1)..];

        try
        {
            int hresult = NativeMethods.TaskDialog(
                IntPtr.Zero, IntPtr.Zero, Caption, encabezado, cuerpo,
                taskDialogButtons, taskDialogIcon, out int boton);

            if (hresult == 0)
            {
                return boton;
            }
        }
        catch (DllNotFoundException)
        {
        }
        catch (EntryPointNotFoundException)
        {
        }

        return NativeMethods.MessageBoxW(
            IntPtr.Zero, message, Caption,
            messageBoxButtons | messageBoxIcon | NativeMethods.MbSetForeground);
    }
}
