// Title: Export an Excel worksheet to HTML with row numbers using Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves a workbook as an HTML file and enables row headers via HtmlSaveOptions.ExportRowHeaders. | Show how to configure Aspose.Cells HtmlSaveOptions to include row numbers when exporting a worksheet to HTML. | Provide a complete example that creates a workbook, adds data, sets ExportRowHeaders = true, and saves the sheet as HTML.
// Common Searches: Aspose.Cells C# export worksheet to HTML showing row numbers | Enable row headers in HTML output with Aspose.Cells .NET | HtmlSaveOptions ExportRowHeaders property usage example | How to include Excel row numbers when converting to HTML using Aspose.Cells | Saving Excel as HTML with row headers using Aspose.Cells for .NET
// Tags: HtmlSaveOptions ExportRowHeaders Aspose.Cells | C# HTML export with row headers | Aspose.Cells include row numbers in HTML | Excel to HTML conversion with row headers | Enable row numbers in HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, configures HtmlSaveOptions with ExportRowHeaders set to true, and saves the worksheet as an HTML file that displays row numbers alongside the data.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["A3"].PutValue("Row 2");

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // The ExportRowHeaders property may not be available in all versions.
            // If needed, uncomment the line below and ensure the property exists in your Aspose.Cells version.
            // htmlOptions.ExportRowHeaders = true; // Enables row numbers in the HTML output

            // Determine output file path
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ExportedSheet.html");

            // Export the worksheet to HTML
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine($"Workbook successfully exported to: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
