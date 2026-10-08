// Title: Export an Aspose.Cells workbook to HTML with a custom TableCssId using HtmlSaveOptions in C#
// AI Prompts: Write C# code that creates a Workbook, sets HtmlSaveOptions.TableCssId to a custom value, and saves the file as HTML. | Show how to apply the same HtmlSaveOptions instance with a set TableCssId to several workbooks for batch HTML export. | Adapt the sample to load an existing Excel file, assign a custom TableCssId, and generate the HTML output.
// Common Searches: Aspose.Cells C# how to set TableCssId in HtmlSaveOptions when saving to HTML | save Excel workbook as HTML with specific table id using Aspose.Cells .NET | pre‑configure HtmlSaveOptions for repeated HTML exports in Aspose.Cells | custom CSS identifier for HTML table generated from Excel with Aspose.Cells | example of Workbook.Save with HtmlSaveOptions.TableCssId in C#
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | C# export workbook to HTML with custom table id | preconfigured HtmlSaveOptions for batch HTML export | custom CSS id for generated HTML table Aspose.Cells | Excel to HTML conversion with table CSS identifier

using Aspose.Cells;
using System;

// // This program creates a new workbook, adds sample data to the first worksheet, configures HtmlSaveOptions with TableCssId set to "myCustomTable", and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Configure HTML save options with a custom TableCssId
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.TableCssId = "myCustomTable";

        // Save the workbook as HTML using the pre‑configured options
        workbook.Save("output.html", htmlOptions);
    }
}
