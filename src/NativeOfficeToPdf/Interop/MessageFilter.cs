using System.Runtime.InteropServices;

namespace NativeOfficeToPdf.Interop;

/// <summary>
/// Filtro de mensajes OLE para el hilo STA que habla con Office.
/// <para>
/// Cuando Word o PowerPoint están ocupados —exportando otro documento, abriendo un archivo de OneDrive
/// que todavía se descarga, o con un diálogo abierto en la instancia del usuario— rechazan las llamadas
/// entrantes con <c>SERVERCALL_RETRYLATER</c>. Sin filtro, COM lo convierte de inmediato en
/// <c>RPC_E_CALL_REJECTED</c> (0x80010001) y la conversión falla. Con este filtro, COM reintenta la
/// llamada cada <see cref="RetryDelayMilliseconds"/> hasta <see cref="GiveUpAfterMilliseconds"/>. Es la
/// solución que documenta Microsoft para automatizar Office desde un proceso externo.
/// </para>
/// </summary>
// ComVisible explícito: COM tiene que poder llamar al filtro a través de IOleMessageFilter.
[ComVisible(true)]
internal sealed class MessageFilter : IOleMessageFilter, IDisposable
{
    private const int ServerCallIsHandled = 0;
    private const int ServerCallRetryLater = 2;
    private const int PendingMsgWaitDefProcess = 2;
    private const int CancelCall = -1;

    private const int RetryDelayMilliseconds = 250;
    private const int GiveUpAfterMilliseconds = 120_000;

    private readonly IOleMessageFilter? previous;
    private bool disposed;

    private MessageFilter()
    {
        // Si no se puede registrar (hilo MTA, por ejemplo), se sigue sin filtro: la conversión funciona
        // igual mientras Office no esté ocupado, que es el comportamiento anterior.
        Registered = NativeMethods.CoRegisterMessageFilter(this, out previous) == 0;
    }

    private bool Registered { get; }

    /// <summary>Registra el filtro en el hilo actual. <see cref="Dispose"/> repone el anterior.</summary>
    public static MessageFilter Register() => new();

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        if (Registered)
        {
            NativeMethods.CoRegisterMessageFilter(previous, out _);
        }
    }

    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) =>
        ServerCallIsHandled;

    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) =>
        dwRejectType == ServerCallRetryLater && dwTickCount < GiveUpAfterMilliseconds
            ? RetryDelayMilliseconds
            : CancelCall;

    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) =>
        PendingMsgWaitDefProcess;
}

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

    [PreserveSig]
    int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}
