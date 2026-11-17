// Title: Save an Excel workbook as HTML with a custom TableCssId using Aspose.Cells for .NET (C# shortcut)
// Common Searches: C# Aspose.Cells how to assign a TableCssId when exporting to HTML | using HtmlSaveOptions to set a custom table CSS id in Excel to HTML conversion | Aspose.Cells Save method with predefined HtmlSaveOptions for quick HTML export | export workbook to HTML with a specific table id in .NET | Aspose.Cells shortcut for SaveAsHtml with custom TableCssId
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | C# Excel to HTML conversion with custom table id | preconfigured SaveAsHtml shortcut Aspose.Cells | custom table CSS identifier for HTML export | Aspose.Cells HTML export options example

using Aspose.Cells;
using System;

// The example creates a Workbook, adds sample data, configures HtmlSaveOptions with TableCssId set to "myCustomTable", and saves the workbook as an HTML file using the pre‑configured options for a fast, reusable export.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Value");

        // Configure HtmlSaveOptions with a custom TableCssId
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.TableCssId = "myCustomTable";

        // Save the workbook as HTML using the pre‑configured options
        workbook.Save("ExportedWorkbook.html", htmlOptions);
    }
}
