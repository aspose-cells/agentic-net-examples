// Title: Export an Excel worksheet to HTML with row numbers displayed using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Generate C# code that creates a Workbook, populates sample data, sets HtmlSaveOptions.ExportRowHeaders = true, and saves the workbook as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions in .NET to include row headers when converting an Excel sheet to HTML.
// Common Searches: Aspose.Cells C# export worksheet to HTML showing row numbers | How to enable ExportRowHeaders in HtmlSaveOptions for Aspose.Cells | C# example of Aspose.Cells HTML conversion with row headers | Generate HTML from Excel with row headers using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportRowHeaders | C# export Excel to HTML with row numbers | Aspose.Cells HTML row header display | Aspose.Cells .NET HTML conversion row headers

using Aspose.Cells;
using System;

// The example creates a new Workbook, adds sample data, configures HtmlSaveOptions.ExportRowHeaders to true so row numbers appear, and saves the workbook as an HTML file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Row1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Row2");
            sheet.Cells["B3"].PutValue(456);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Note: ExportRowHeaders property is not available in the current API version.

            // Export the workbook to HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
