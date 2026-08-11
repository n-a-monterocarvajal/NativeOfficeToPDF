using System.Runtime.InteropServices;
using NativeOfficeToPdf.Interop;

namespace NativeOfficeToPdf.Converters;

/// <summary>
/// PowerPoint → PDF con <c>Presentation.ExportAsFixedFormat</c>.
/// <para>
/// A diferencia de Word, la ventana no se oculta con <c>Visible = false</c> — PowerPoint rechaza esa
/// asignación. La vía soportada es abrir la presentación con <c>WithWindow:=msoFalse</c>.
/// </para>
/// </summary>
internal sealed class PowerPointConverter : IDocumentConverter
{
    private const string ProgId = "PowerPoint.Application";

    // MsoTriState
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    private const int PpAlertsNone = 1;
    private const int PpFixedFormatTypePdf = 2;

    // Intent "Print" en vez de "Screen": es la calidad alta, que es el objetivo de todo el proyecto.
    private const int PpFixedFormatIntentPrint = 2;
    private const int PpPrintOutputSlides = 1;
    private const int PpPrintHandoutVerticalFirst = 1;
    private const int PpPrintAll = 1;

    private const int MsoAutomationSecurityForceDisable = 3;

    public void Convert(string sourcePath, string destinationPath)
    {
        using OfficeApplication powerPoint = OfficeApplication.GetOrCreate(ProgId);
        ComObject application = powerPoint.Application;

        object? previousAlerts = application.TryGetProperty("DisplayAlerts");
        object? previousSecurity = application.TryGetProperty("AutomationSecurity");

        try
        {
            application.TrySetProperty("DisplayAlerts", PpAlertsNone);
            application.TrySetProperty("AutomationSecurity", MsoAutomationSecurityForceDisable);

            using ComObject presentations = application.GetObject("Presentations");
            using ComObject presentation = presentations.InvokeNamedForObject(
                "Open",
                ("FileName", sourcePath),
                ("ReadOnly", MsoTrue),
                ("Untitled", MsoFalse),
                ("WithWindow", MsoFalse));

            try
            {
                presentation.InvokeNamed(
                    "ExportAsFixedFormat",
                    ("Path", destinationPath),
                    ("FixedFormatType", PpFixedFormatTypePdf),
                    ("Intent", PpFixedFormatIntentPrint),
                    ("FrameSlides", MsoFalse),
                    ("HandoutOrder", PpPrintHandoutVerticalFirst),
                    ("OutputType", PpPrintOutputSlides),
                    ("PrintHiddenSlides", MsoFalse),
                    // PrintRange figura como "requerido" en la referencia VBA aunque puede ir a
                    // Nothing: con enlace tardío (sin typelib de por medio) omitirlo hace que Office
                    // no pueda resolver su valor por defecto y devuelva DISP_E_TYPEMISMATCH
                    // (0x80020005) en vez de exportar. Pasar null explícito evita el fallo.
                    ("PrintRange", null),
                    ("RangeType", PpPrintAll),
                    ("IncludeDocProperties", true),
                    ("KeepIRMSettings", true),
                    ("DocStructureTags", true),
                    ("BitmapMissingFonts", true),
                    ("UseISO19005_1", false));
            }
            finally
            {
                presentation.Invoke("Close");
            }
        }
        catch (COMException ex)
        {
            throw new OfficeAutomationException(
                $"PowerPoint no pudo convertir «{Path.GetFileName(sourcePath)}»: {ex.Message}", ex);
        }
        finally
        {
            if (previousSecurity is not null)
            {
                application.TrySetProperty("AutomationSecurity", previousSecurity);
            }

            if (previousAlerts is not null)
            {
                application.TrySetProperty("DisplayAlerts", previousAlerts);
            }
        }
    }
}
