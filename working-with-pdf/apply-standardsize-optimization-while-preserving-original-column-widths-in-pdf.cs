// Title: Keeping Excel column widths unchanged during PDF conversion with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# keep column widths when saving workbook as PDF | PdfSaveOptions AllColumnsInOnePagePerSheet false example in .NET | Convert Excel to PDF without scaling columns using Aspose.Cells | Reduce PDF file size while preserving layout in Aspose.Cells C# | How to prevent column compression in PDF export with Aspose.Cells
// Tags: Aspose.Cells PDF column width preservation | PdfSaveOptions AllColumnsInOnePagePerSheet false | Excel to PDF conversion without column scaling | C# Aspose.Cells PDF size optimization | preserve layout during PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, configures PdfSaveOptions with AllColumnsInOnePagePerSheet set to false to retain the original column widths, and saves the workbook as a PDF, ensuring the layout is unchanged and the PDF size is optimized.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    var sampleWorkbook = new Workbook();
                    var sheet = sampleWorkbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Sample Data");
                    sampleWorkbook.Save(inputPath);
                }

                // Load the workbook.
                var workbook = new Workbook(inputPath);

                // Configure PDF save options.
                var pdfOptions = new PdfSaveOptions
                {
                    // Preserve original column widths by not forcing all columns onto one page.
                    AllColumnsInOnePagePerSheet = false
                };

                // Save the workbook as PDF.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
