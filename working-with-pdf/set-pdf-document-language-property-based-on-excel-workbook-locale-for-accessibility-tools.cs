using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfLanguageDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Determine the workbook's locale.
            // Prefer the CultureInfo set in WorkbookSettings; if not set, fall back to a default.
            string locale = workbook.Settings.CultureInfo != null
                ? workbook.Settings.CultureInfo.Name          // e.g., "en-US", "fr-FR"
                : "en-US";

            // Set the built‑in document property that maps to the PDF language metadata.
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;
            builtInProps.Language = locale;

            // Create PDF save options (optional: set default edit language if needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: keep auto detection; adjust if you need specific language handling
                DefaultEditLanguage = DefaultEditLanguage.Auto,
                // Ensure fonts are checked to avoid missing characters
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF; the language property will be embedded for accessibility tools.
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine($"Workbook saved as PDF with language set to '{locale}'.");
        }
    }
}