// Title: Load an HTML file into an Aspose.Cells workbook and re‑save it with a custom TableCssId using C#
// AI Prompts: Read an existing HTML file into a Workbook object, set HtmlSaveOptions.TableCssId to a custom value, and save the workbook as HTML using Aspose.Cells for .NET. | Use the Aspose.Cells C# API to change the table CSS identifier when converting a workbook loaded from HTML back to HTML.
// Common Searches: Aspose.Cells C# load html file into workbook and change table id on export | set TableCssId in HtmlSaveOptions when saving workbook as HTML Aspose.Cells | how to modify HTML table CSS id using Aspose.Cells .NET | re‑export HTML to HTML with different TableCssId using Aspose.Cells | C# Aspose.Cells example for loading HTML and customizing table CSS identifier
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | load HTML into Workbook C# | export workbook to HTML with custom table id | change HTML table CSS identifier Aspose.Cells | C# Aspose.Cells HTML conversion custom TableCssId

using System;
using Aspose.Cells;

// // Loads an HTML file into an Aspose.Cells Workbook, sets HtmlSaveOptions.TableCssId to "newTableId", and saves the workbook back to HTML.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string inputPath = "input.html";

        // Load the HTML file into a workbook
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options with a new TableCssId
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.TableCssId = "newTableId";

        // Path for the exported HTML file
        string outputPath = "output.html";

        // Export the workbook to HTML using the specified options
        workbook.Save(outputPath, htmlOptions);
    }
}
