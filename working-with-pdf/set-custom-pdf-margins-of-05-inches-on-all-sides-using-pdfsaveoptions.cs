// Title: How to set 0.5‑inch page margins for every worksheet and export to PDF with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells set worksheet margins to half an inch before PDF conversion | Aspose.Cells PdfSaveOptions custom margins example in .NET | Set all worksheet page margins to 0.5 inches using Aspose.Cells API | Export Excel to PDF with 0.5 inch margins using Aspose.Cells C# | Adjust PDF page margins in Aspose.Cells workbook export
// Tags: worksheet margin setup Aspose.Cells | PdfSaveOptions margin customization C# | half‑inch page margins PDF export | Aspose.Cells export Excel to PDF with custom margins | C# set page margins inches Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The code loads or creates an Excel workbook, applies a 0.5‑inch top, bottom, left, and right margin to each worksheet via PageSetup, configures PdfSaveOptions, and saves the workbook as a PDF file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook or load an existing one if the file is present.
                Workbook workbook;
                string inputPath = "input.xlsx";

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with a default worksheet
                }

                // Set 0.5‑inch margins for all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.TopMargin = 0.5;    // inches
                    sheet.PageSetup.BottomMargin = 0.5;
                    sheet.PageSetup.LeftMargin = 0.5;
                    sheet.PageSetup.RightMargin = 0.5;
                }

                // Configure PDF save options (default options are sufficient for this example).
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF.
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
