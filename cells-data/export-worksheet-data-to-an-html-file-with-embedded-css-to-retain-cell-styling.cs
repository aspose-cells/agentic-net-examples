using System;
using System.Drawing;
using Aspose.Cells;

namespace ExportWorksheetToHtml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Apply simple styling to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Configure HTML save options to embed CSS (default behavior)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Ensure CSS is embedded in the HTML file
                ExportWorksheetCSSSeparately = false,
                // Keep CSS usage (do not force inline styles only)
                DisableCss = false,
                // Optional: add custom CSS for the whole page
                CssStyles = @"
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    table { border-collapse: collapse; }
                    td, th { border: 1px solid #ccc; padding: 5px 10px; }
                "
            };

            // Define output HTML file path
            string outputPath = "WorksheetExport.html";

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Worksheet exported to HTML with embedded CSS at: {outputPath}");
        }
    }
}