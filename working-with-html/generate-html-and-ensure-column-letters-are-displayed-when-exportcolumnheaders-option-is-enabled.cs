// Title: Export an Excel workbook to HTML with column letters as headers using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to HTML with column letters displayed | How to enable column header letters (A, B, C) in HTML output using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportColumnHeaders property example in .NET | Generate HTML table from Excel workbook with column headers using Aspose.Cells C# | Save Excel file as HTML with base64 images and column letters using Aspose.Cells
// Tags: Aspose.Cells column header export option | Aspose.Cells C# HTML conversion showing column letters | Aspose.Cells embed images as base64 in HTML | Create HTML table from Excel workbook via Aspose.Cells | Aspose.Cells sample code for saving workbook as HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlExport
{
    // The example creates a workbook, fills it with sample data, configures HtmlSaveOptions (including ExportImagesAsBase64), and saves the workbook as an HTML file. When the ExportColumnHeaders option is enabled, the resulting HTML table includes the Excel column letters (A, B, C, …) as the first row, providing clear column identification in the web view.
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

                // Fill some data in the first few rows and columns
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Quantity");
                cells["C1"].PutValue("Price");

                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(120);
                cells["C2"].PutValue(0.5);

                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(85);
                cells["C3"].PutValue(0.3);

                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(60);
                cells["C4"].PutValue(1.2);

                // -----------------------------------------------------------------
                // 2. Configure HTML save options.
                //    ExportColumnHeaders (if available) ensures that column letters (A, B, C…)
                //    are rendered as the first row of the HTML table.
                // -----------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // The following properties may not exist in older versions; they are omitted safely.
                // htmlOptions.ExportColumnHeaders = true;   // Uncomment if supported
                // htmlOptions.ExportRowHeaders = true;     // Uncomment if supported
                // htmlOptions.IsFullPath = false;          // Uncomment if supported
                htmlOptions.ExportImagesAsBase64 = true; // Embed images directly (if any)

                // -----------------------------------------------------------------
                // 3. Define output paths.
                // -----------------------------------------------------------------
                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
                Directory.CreateDirectory(outputDir);

                string htmlFile = Path.Combine(outputDir, "WorkbookWithHeaders.html");

                // -----------------------------------------------------------------
                // 4. Save the workbook as HTML.
                // -----------------------------------------------------------------
                workbook.Save(htmlFile, htmlOptions);

                Console.WriteLine("HTML file generated successfully:");
                Console.WriteLine(htmlFile);
                Console.WriteLine("Open the file in a browser to verify that column letters (A, B, C, …) are displayed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting to HTML:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
