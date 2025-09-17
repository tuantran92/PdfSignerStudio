using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

using iText.Forms;
using iText.Forms.Fields;
using iText.Forms.Fields.Properties;
using iText.Kernel.Exceptions;
using iText.Kernel.Pdf;
// KHÔNG import iText.Kernel.Geom.* để tránh đụng tên Path.
// Alias Rectangle của iText:
using PdfRect = iText.Kernel.Geom.Rectangle;

namespace PdfSignerStudio
{
    public static class PdfService
    {
        public static string ConvertDocxToPdfWithWord(string docxPath, string outDir)
        {
            Directory.CreateDirectory(outDir);
            var pdfPath = System.IO.Path.Combine(outDir, System.IO.Path.GetFileNameWithoutExtension(docxPath) + ".pdf");

            Word.Application? app = null;
            Word.Document? doc = null;

            try
            {
                app = new Word.Application
                {
                    Visible = false,
                    ScreenUpdating = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };

                object missing = Type.Missing;
                object readOnly = true;
                object addToRecent = false;
                object isVisible = false;
                object no = false;
                object fileObj = docxPath;

                doc = app.Documents.Open(
                    ref fileObj,
                    ref no,
                    ref readOnly,
                    ref addToRecent,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref isVisible,
                    ref missing, ref missing, ref missing, ref missing
                );

                doc.ExportAsFixedFormat(
                    OutputFileName: pdfPath,
                    ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false,
                    OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range: Word.WdExportRange.wdExportAllDocument,
                    From: 0, To: 0,
                    Item: Word.WdExportItem.wdExportDocumentContent,
                    IncludeDocProps: true,
                    KeepIRM: true,
                    CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks,
                    DocStructureTags: true,
                    BitmapMissingFonts: true,
                    UseISO19005_1: false
                );

                if (!File.Exists(pdfPath))
                    throw new Exception("Word Interop export failed: output PDF not found.");

                return pdfPath;
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(false); } catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }
                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                }
                doc = null; app = null;
                GC.Collect(); GC.WaitForPendingFinalizers();
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }

        public static void AddSignatureFields(string inputPdf, string outputPdf, IEnumerable<FormFieldDef> defs)
        {
            var workIn = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PdfSignerStudio", "work-input.pdf");
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(workIn)!);
            File.Copy(inputPdf, workIn, true);

            try
            {
                using var pdf = new PdfDocument(new PdfReader(workIn), new PdfWriter(outputPdf));
                var form = PdfAcroForm.GetAcroForm(pdf, true);

                var existing = form.GetAllFormFields().Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);

                foreach (var d in defs.Where(x => x.Type == "signature"))
                {
                    if (d.Page < 1 || d.Page > pdf.GetNumberOfPages()) continue;

                    var page = pdf.GetPage(d.Page);
                    PdfRect crop = page.GetCropBox();
                    if (crop == null)
                    {
                        // media box của iText cũng là PdfRect
                        PdfRect mb = page.GetMediaBox();
                        crop = new PdfRect(mb.GetX(), mb.GetY(), mb.GetWidth(), mb.GetHeight());
                    }

                    // clamp vào trang
                    float x = Math.Max(0, d.Rect.X);
                    float y = Math.Max(0, d.Rect.Y);
                    float w = Math.Max(0, d.Rect.W);
                    float h = Math.Max(0, d.Rect.H);

                    float maxW = crop.GetWidth() - x;
                    float maxH = crop.GetHeight() - y;
                    if (maxW <= 0 || maxH <= 0) continue;

                    w = MathF.Min(w, maxW);
                    h = MathF.Min(h, maxH);
                    if (w < 1f || h < 1f) continue;

                    var rect = new PdfRect(x, y, w, h);

                    // tên duy nhất
                    string baseName = string.IsNullOrWhiteSpace(d.Name) ? "Signature" : d.Name.Trim();
                    string name = baseName;
                    int idx = 1;
                    while (existing.Contains(name))
                        name = $"{baseName}_{idx++}";
                    existing.Add(name);

                    var sig = new iText.Forms.Fields.SignatureFormFieldBuilder(pdf, name)
                                  .SetWidgetRectangle(rect)
                                  .CreateSignature();

                    if (d.Required) sig.SetRequired(true);
                    form.AddField(sig, page);
                }

                pdf.Close();
            }
            catch (iText.Kernel.Exceptions.PdfException ex)
            {
                throw new Exception($"iText PdfException: {ex.Message}", ex);
            }
        }
    }
}
