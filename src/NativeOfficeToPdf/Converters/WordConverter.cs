using System.Runtime.InteropServices;
using NativeOfficeToPdf.Interop;

namespace NativeOfficeToPdf.Converters;

/// <summary>
/// Word → PDF con <c>Document.ExportAsFixedFormat</c>, el mismo motor que "Guardar como PDF" de la
/// cinta. No se usa <c>SaveAs2</c>: es más genérico y da menos control sobre marcadores, propiedades
/// del documento e IRM.
/// </summary>
internal static class WordConverter
{
    private const string ProgId = "Word.Application";

    // Constantes de la API de Word (WdSaveOptions, WdExportFormat, …). Con enlace tardío no hay
    // enumeraciones importadas, así que se declaran acá con el nombre que tienen en la referencia VBA.
    private const int WdAlertsNone = 0;
    private const int WdDoNotSaveChanges = 0;
    private const int WdExportFormatPdf = 17;
    private const int WdExportOptimizeForPrint = 0;
    private const int WdExportAllDocument = 0;
    private const int WdExportDocumentContent = 0;
    private const int WdExportCreateHeadingBookmarks = 1;

    // Contraseña deliberadamente inverosímil. Si el documento está protegido, Word falla con una
    // excepción en vez de abrir un diálogo modal que dejaría el proceso colgado sin interfaz.
    private const string ImposiblePassword = " NativeOfficeToPdf-sin-contrasena ";

    public static void Convert(string sourcePath, string destinationPath)
    {
        using OfficeApplication word = OfficeApplication.GetOrCreate(ProgId, WdAlertsNone, WdDoNotSaveChanges);
        ComObject application = word.Application;

        try
        {
            // Solo se oculta la ventana si la instancia es nuestra: si es la del usuario, esconderle
            // Word mientras trabaja sería peor que el problema que se intenta evitar.
            if (word.StartedByUs)
            {
                application.TrySetProperty("Visible", false);
            }

            using ComObject documents = application.GetObject("Documents");
            using ComObject document = documents.InvokeNamedForObject(
                "Open",
                ("FileName", sourcePath),
                ("ConfirmConversions", false),
                ("ReadOnly", true),
                ("AddToRecentFiles", false),
                ("PasswordDocument", ImposiblePassword),
                ("PasswordTemplate", ImposiblePassword),
                ("Revert", false),
                ("Visible", false),
                // En true: si Word detecta contenido no legible, lo repara igual que si el usuario
                // respondiera "Sí" en el diálogo de recuperación, en vez de dejarlo pendiente de una
                // interacción que en un proceso automatizado (Visible=false) nunca llega.
                ("OpenAndRepair", true),
                ("NoEncodingDialog", true));

            try
            {
                document.InvokeNamed(
                    "ExportAsFixedFormat",
                    ("OutputFileName", destinationPath),
                    ("ExportFormat", WdExportFormatPdf),
                    ("OpenAfterExport", false),
                    ("OptimizeFor", WdExportOptimizeForPrint),
                    ("Range", WdExportAllDocument),
                    ("Item", WdExportDocumentContent),
                    ("IncludeDocProps", true),
                    ("KeepIRM", true),
                    ("CreateBookmarks", WdExportCreateHeadingBookmarks),
                    ("DocStructureTags", true),
                    ("BitmapMissingFonts", true),
                    ("UseISO19005_1", false));
            }
            finally
            {
                document.Invoke("Close", WdDoNotSaveChanges);
            }
        }
        catch (COMException ex)
        {
            throw new OfficeAutomationException(
                $"Word no pudo convertir «{Path.GetFileName(sourcePath)}»: {ex.Message}", ex);
        }
    }
}
