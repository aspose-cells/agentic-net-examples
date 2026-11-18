// Title: Apply 0.5‑inch margins to every worksheet and export to PDF using Aspose.Cells PdfSaveOptions in C#
// Common Searches: C# code to set 0.5 inch page margins before saving Excel as PDF with Aspose.Cells | How to apply uniform margins to all worksheets when exporting to PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions example for custom page margins in C# | Setting top, bottom, left, right margins in Excel workbook for PDF output via Aspose.Cells
// Tags: Aspose.Cells worksheet page margins points | PdfSaveOptions configure PDF page layout C# | Excel workbook margin adjustment for PDF export | Uniform half‑inch margins with Aspose.Cells | C# Aspose.Cells PDF conversion margin settings

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The sample loads or creates a workbook, sets each worksheet's top, bottom, left, and right margins to 36 points (0.5 inches), configures default PdfSaveOptions, and saves the workbook as a PDF file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input workbook path (optional)
                string inputPath = "input.xlsx";

                // Load existing workbook if the file exists; otherwise create a new empty workbook
                Workbook workbook = File.Exists(inputPath)
                    ? new Workbook(inputPath)
                    : new Workbook();

                // Set 0.5‑inch margins (0.5 inch = 36 points) for each worksheet
                const double marginPoints = 36.0;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.TopMargin = marginPoints;
                    sheet.PageSetup.BottomMargin = marginPoints;
                    sheet.PageSetup.LeftMargin = marginPoints;
                    sheet.PageSetup.RightMargin = marginPoints;
                }

                // Configure PDF save options (default options are sufficient for this example)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the custom margin settings
                string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
