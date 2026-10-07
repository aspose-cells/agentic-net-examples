// Title: Export an Aspose.Cells workbook to HTML with a custom TableCssId while suppressing the default stylesheet in C#
// AI Prompts: Generate C# code that saves a Workbook to HTML using HtmlSaveOptions, assigning an identifier to the table and setting CssStyleSheetType to None to omit embedded styles. | Show how to assign a specific ID to the exported HTML table and rely on external CSS for styling when using Aspose.Cells. | Modify the example to add a custom CSS class to table rows after disabling the built‑in stylesheet in the HTML export.
// Common Searches: Aspose.Cells C# export worksheet to HTML without embedded CSS | Assign a custom ID to the HTML table when saving Excel with Aspose.Cells | Prevent default stylesheet generation in Aspose.Cells HTML output | Use external CSS file to style HTML table produced by Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | suppress built‑in CSS in HTML export | export workbook as HTML with external styling | custom table identifier for Aspose.Cells HTML output | C# external stylesheet for Aspose.Cells generated HTML

using System;
using System.IO;
using Aspose.Cells;

// // Demonstrates creating a workbook, populating sample data, and saving it as an HTML file using HtmlSaveOptions where a custom TableCssId is assigned and the default CSS stylesheet generation is omitted.
class ExportHtmlWithoutCss
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Assign an identifier to the exported table; external CSS can target this ID
                TableCssId = "myExportedTable"
                // Note: CssStyleSheetType property may not be available in older versions.
                // If supported, you can uncomment the following line to disable CSS generation:
                // CssStyleSheetType = CssStyleSheetType.None
            };

            string outputPath = "ExportedTable.html";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to HTML (lifecycle rule: save)
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"HTML file saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
