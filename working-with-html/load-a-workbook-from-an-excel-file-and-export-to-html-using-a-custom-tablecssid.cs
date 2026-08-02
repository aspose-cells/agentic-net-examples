// Title: C# – Export Excel to HTML with a Custom TableCssId using Aspose.Cells
// Description: Load an .xlsx workbook, set HtmlSaveOptions.TableCssId to a custom value, and save the file as HTML. The generated <table> element receives the specified CSS ID for easy styling and scripting.
// Keywords: Aspose.Cells C# HTML export | TableCssId property | custom CSS ID for exported table | HtmlSaveOptions SaveFormat.Html | .NET Excel to HTML conversion | style Aspose.Cells HTML output
// Common Searches: Aspose.Cells set TableCssId when saving as HTML | C# export Excel workbook to HTML with custom table id | HtmlSaveOptions TableCssId example | how to assign CSS ID to HTML table using Aspose.Cells | customize HTML table id in Aspose.Cells export
// Developer Intent: Create an HTML file from an Excel workbook and assign a unique CSS ID to the resulting table element.
// Use Cases: Apply external stylesheet rules to the exported table via its custom ID. | Enable JavaScript to locate and manipulate the table in a web page. | Generate multiple reports where each HTML table has a distinct identifier for individualized styling.
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as HTML, setting TableCssId to "my-table". | Show how to combine TableCssId with other HtmlSaveOptions like ExportImagesAsBase64 and CssStyleSheetType. | Explain how to reference the custom TableCssId in an external CSS file to style the exported HTML table.

using System;
using Aspose.Cells;

// Load an .xlsx workbook, set HtmlSaveOptions.TableCssId to a custom value, and save the file as HTML. The generated <table> element receives the specified CSS ID for easy styling and scripting.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and set a custom TableCssId
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "custom-table-style";

        // Save the workbook as HTML using the configured options
        string outputPath = "output.html";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"HTML file saved with TableCssId = {saveOptions.TableCssId}");
    }
}
