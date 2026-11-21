// Title: Save an Excel workbook to PDF with charts retained as vector graphics using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF keeping chart vectors | How to prevent chart rasterization when converting XLSX to PDF with Aspose.Cells | Preserve vector quality of Excel charts in PDF output using Aspose.Cells .NET | C# PdfSaveOptions to retain chart scalability in Excel to PDF conversion | Generate PDF from Excel workbook with vector charts using Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions chart vector rendering | C# Excel workbook PDF export preserving vector graphics | vector chart export Aspose.Cells .NET | high‑resolution chart PDF conversion Aspose.Cells | Excel chart scalability PDF output C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

namespace AsposeCellsPdfConversion
{
    // The example loads an Excel file (or creates a sample workbook with a column chart), configures PdfSaveOptions (which keep charts as vectors by default), and saves the workbook as a PDF, ensuring chart scalability and clarity in the resulting document.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Ensure the input file exists; create a simple workbook with a chart if it does not.
                if (!File.Exists(inputPath))
                {
                    CreateSampleWorkbook(inputPath);
                    Console.WriteLine($"Sample workbook created at '{inputPath}'.");
                }

                // Load the existing Excel workbook.
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default behavior keeps charts as vectors).
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF using the configured options.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void CreateSampleWorkbook(string path)
        {
            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Name = "Data";

                // Populate sample data.
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B4"].PutValue(30);

                // Add a column chart.
                int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 7);
                Chart chart = ws.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Sample Chart";

                // Save the workbook to the specified path.
                wb.Save(path);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to create sample workbook: {ex.Message}");
                throw; // Re‑throw to let the caller handle it.
            }
        }
    }
}
