// Title: Generate HTML from an Excel workbook with visible row numbers using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export workbook to HTML showing row numbers | How to enable ExportRowColumnHeadings in HtmlSaveOptions for HTML output | Aspose.Cells HtmlSaveOptions PresentationPreference with row headers example | Save Excel as HTML with row and column headings using Aspose.Cells .NET | Display row numbers in HTML generated from Excel using Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions ExportRowColumnHeadings | HTML export with row numbers Aspose.Cells | C# Aspose.Cells save workbook as HTML | PresentationPreference option Aspose.Cells | Excel to HTML conversion with row headers

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // The sample creates a Workbook, fills it with product data, configures HtmlSaveOptions to set ExportRowColumnHeadings=true and PresentationPreference=true, and saves the file as HTML. The resulting HTML displays row numbers on the left side of the grid.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and populate it with sample data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add header row
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Quantity");
                cells["C1"].PutValue("Price");

                // Add some sample rows
                string[,] data = new string[,]
                {
                    { "Apple",  "120", "0.50" },
                    { "Banana", "80",  "0.30" },
                    { "Cherry", "150", "0.20" },
                    { "Date",   "60",  "1.00" }
                };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    cells[i + 1, 0].PutValue(data[i, 0]); // Product
                    cells[i + 1, 1].PutValue(int.Parse(data[i, 1])); // Quantity
                    cells[i + 1, 2].PutValue(double.Parse(data[i, 2])); // Price
                }

                // -----------------------------------------------------------------
                // 2. Configure HTML save options to export row headers (row numbers).
                // -----------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Show row numbers and column headings
                    ExportRowColumnHeadings = true,
                    // Optional: nicer visual layout
                    PresentationPreference = true
                };

                // Define output folder and file name
                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
                Directory.CreateDirectory(outputDir);
                string htmlPath = Path.Combine(outputDir, "WorkbookWithRowHeaders.html");

                // -----------------------------------------------------------------
                // 3. Save the workbook as HTML using the configured options.
                // -----------------------------------------------------------------
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine("HTML file generated successfully:");
                Console.WriteLine(htmlPath);
                Console.WriteLine();
                Console.WriteLine("Open the HTML file in a browser to verify that row numbers are displayed on the left side of the grid.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the HTML export process:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
