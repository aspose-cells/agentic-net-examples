// Title: Export an Excel workbook to HTML with a custom TableCssId and verify no CSS class name collisions using Aspose.Cells for .NET (C#)
// Common Searches: how to assign a custom TableCssId in Aspose.Cells HtmlSaveOptions C# | detect CSS class name conflicts after exporting a workbook to HTML with Aspose.Cells | verify that the TableCssId does not appear as a CSS class selector in generated HTML | C# example for checking style collisions in Aspose.Cells HTML export | prevent id and class naming conflicts when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | custom table identifier for Excel to HTML conversion | detect CSS class collisions in generated HTML | C# validation of HTML export style conflicts | prevent id/class naming conflicts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills it with data, configures HtmlSaveOptions to set a custom TableCssId, saves the workbook as HTML to a memory stream, reads the HTML content, checks for any CSS class selector matching the TableCssId to detect naming collisions, throws an exception if a conflict is found, writes the HTML to a file, and logs a success message.
class HtmlExportWithCustomTableId
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header 1");
        sheet.Cells["B1"].PutValue("Header 2");
        sheet.Cells["A2"].PutValue("Row 1 Col 1");
        sheet.Cells["B2"].PutValue("Row 1 Col 2");
        sheet.Cells["A3"].PutValue("Row 2 Col 1");
        sheet.Cells["B3"].PutValue("Row 2 Col 2");

        // Configure HTML save options with a custom TableCssId
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Assign a unique ID for the generated HTML table
            TableCssId = "myCustomTable"
        };

        // Save the workbook to a memory stream as HTML
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);
            htmlStream.Position = 0;

            // Read the generated HTML into a string
            string htmlContent;
            using (StreamReader reader = new StreamReader(htmlStream))
            {
                htmlContent = reader.ReadToEnd();
            }

            // Verify that the custom TableCssId does not collide with any CSS class names
            // A collision would appear as a CSS class selector like ".myCustomTable"
            if (htmlContent.Contains(".myCustomTable"))
            {
                throw new InvalidOperationException(
                    "CSS class name collision detected: the TableCssId 'myCustomTable' is also used as a CSS class.");
            }

            // Optionally, write the HTML to a file for inspection
            File.WriteAllText("ExportedWorkbook.html", htmlContent);
        }

        Console.WriteLine("HTML export completed successfully with TableCssId='myCustomTable'.");
    }
}
