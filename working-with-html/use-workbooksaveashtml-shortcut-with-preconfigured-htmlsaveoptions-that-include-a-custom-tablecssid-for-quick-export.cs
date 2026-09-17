// Title: Save an Aspose.Cells workbook as HTML with a custom TableCssId using HtmlSaveOptions in C#
// AI Prompts: Write C# code that creates a workbook, populates cells, configures the HTML export settings to assign a custom CSS identifier to the generated table, and saves the workbook as an HTML file. | Show how to export a workbook to HTML and give the resulting table a custom CSS ID by adjusting the HTML export options before invoking the Save method in C#.
// Common Searches: Aspose.Cells C# export workbook to HTML with custom table identifier | How to set a CSS ID for the HTML table when converting Excel to HTML using Aspose.Cells | C# example for saving Excel as HTML with a specific table id | Assign custom table CSS id during Aspose.Cells HTML conversion
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | C# save workbook as HTML | custom table CSS id Aspose.Cells | HTML export options Aspose.Cells | Workbook.Save HtmlSaveOptions example

using System;
using Aspose.Cells;

// Creates a workbook, fills it with sample data, sets HtmlSaveOptions.TableCssId to "myCustomTable", and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Row 2");
            sheet.Cells["B3"].PutValue(200);

            // Configure HtmlSaveOptions with a custom TableCssId
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.TableCssId = "myCustomTable";

            // Save the workbook as HTML using the correct Save method
            workbook.Save("ExportedWorkbook.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
