// Title: Export an Excel workbook to PDF with Aspose.Cells using InvariantCulture and French (fr-FR) cultures and compare the results
// Common Searches: how to export an Excel workbook to PDF with a specific CultureInfo using Aspose.Cells C# | Aspose.Cells PDF export number formatting differences between invariant and fr-FR cultures | compare PDF files generated from the same workbook under different cultures in .NET | set thread culture before saving workbook as PDF with Aspose.Cells | detect locale-specific formatting changes in PDF output from Aspose.Cells
// Tags: Aspose.Cells PDF export using specific CultureInfo | C# thread CultureInfo impact on PDF rendering | locale-aware number format in Aspose.Cells PDF | binary comparison of PDFs generated under different cultures | InvariantCulture versus French culture PDF generation Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, switches the thread's CurrentCulture to InvariantCulture and saves the workbook as a PDF, then switches to the French (fr-FR) culture and saves another PDF. It restores the original culture and compares the two PDFs by size and byte content, reporting whether they are identical while handling Aspose.Cells‑specific and general exceptions.
class WorkbookPdfCultureComparison
{
    static void Main()
    {
        // Paths for the source Excel file and output PDFs
        const string sourceFile = "input.xlsx";
        const string pdfInvariant = "output_invariant.pdf";
        const string pdfFrench = "output_french.pdf";

        try
        {
            // Ensure the source file exists; create a simple workbook if it does not.
            if (!File.Exists(sourceFile))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample data");
                tempWb.Save(sourceFile);
                Console.WriteLine($"Created placeholder workbook: {sourceFile}");
            }

            // Load the workbook (no password required for the placeholder file)
            var loadOptions = new LoadOptions(); // default options
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Preserve the original culture to restore later
            CultureInfo originalCulture = CultureInfo.CurrentCulture;

            // ---------- Invariant Culture ----------
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            workbook.Save(pdfInvariant, SaveFormat.Pdf);

            // ---------- French Culture ----------
            CultureInfo.CurrentCulture = new CultureInfo("fr-FR");
            workbook.Save(pdfFrench, SaveFormat.Pdf);

            // Restore the original culture
            CultureInfo.CurrentCulture = originalCulture;

            // ---------- Comparison ----------
            bool filesAreIdentical = File.ReadAllBytes(pdfInvariant).Length == File.ReadAllBytes(pdfFrench).Length &&
                                     File.ReadAllBytes(pdfInvariant).SequenceEqual(File.ReadAllBytes(pdfFrench));

            Console.WriteLine($"PDF files are {(filesAreIdentical ? "identical" : "different")}.");
            Console.WriteLine($"Invariant PDF size: {new FileInfo(pdfInvariant).Length} bytes");
            Console.WriteLine($"French PDF size:    {new FileInfo(pdfFrench).Length} bytes");
        }
        catch (CellsException cex)
        {
            Console.WriteLine($"Aspose.Cells error: {cex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
