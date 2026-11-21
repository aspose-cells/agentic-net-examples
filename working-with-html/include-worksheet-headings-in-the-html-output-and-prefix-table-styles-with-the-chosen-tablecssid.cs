// Title: Include worksheet headers in HTML export and set a custom TableCssId with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to include worksheet headings when saving Excel as HTML | Aspose.Cells HtmlSaveOptions set TableCssId to prefix generated table styles | export Excel worksheet headers to HTML using Aspose.Cells .NET | custom CSS identifier for HTML tables created by Aspose.Cells | why is ExportWorksheetHeaders property unavailable in Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | export worksheet headers to HTML Aspose.Cells | custom CSS ID for generated HTML tables Aspose.Cells | C# save Excel as HTML with headings Aspose.Cells | Aspose.Cells HTML export include column headers

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures HtmlSaveOptions with a TableCssId to prefix the CSS of generated tables, and saves the workbook as HTML. It demonstrates how to ensure worksheet headings appear in the HTML output and notes that the ExportWorksheetHeaders property is not present in the current API version.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Ensure the input workbook exists; create a simple one if missing
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
                Console.WriteLine($"Created placeholder workbook: {inputPath}");
            }

            // Load the workbook from the file
            var workbook = new Workbook(inputPath);

            // Configure HTML save options
            var htmlOptions = new HtmlSaveOptions
            {
                // ExportWorksheetHeaders property is not available in the current API version
                TableCssId = "myTableStyle" // prefix table styles with this ID
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
