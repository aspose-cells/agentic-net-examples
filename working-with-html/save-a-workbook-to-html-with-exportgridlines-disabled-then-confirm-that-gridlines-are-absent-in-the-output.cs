// Title: Save an Aspose.Cells workbook as HTML without gridlines and verify the output
// AI Prompts: Generate C# code that creates a workbook, sets HtmlSaveOptions.ExportGridLines to false, and saves it as an HTML file. | Add C# logic to read the saved HTML file, detect any CSS border definitions, and output a clear status message about gridline presence. | Adapt the example to accept a custom output path via a command‑line argument and report whether gridlines were successfully removed.
// Common Searches: Aspose.Cells how to export Excel to HTML without showing gridlines in .NET | C# check generated HTML from Aspose.Cells for cell border CSS | disable gridlines in HTML output using HtmlSaveOptions ExportGridLines false | verify Aspose.Cells HTML export does not include border styles
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines false | disable gridlines in HTML export C# | verify HTML cell borders Aspose.Cells | save workbook as HTML without borders | read HTML file to detect CSS border styles C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, populates it with data, configures HtmlSaveOptions with ExportGridLines set to false, saves the workbook as an HTML file, reads the generated HTML, checks for CSS border definitions to infer gridlines, and prints a message indicating whether gridlines are absent.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Configure HTML save options to disable gridlines
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.ExportGridLines = false; // Disable gridlines

        // Define output HTML file path
        string htmlPath = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlPath, saveOptions);

        // Verify that gridlines are absent in the generated HTML
        string htmlContent = File.ReadAllText(htmlPath);

        // Simple check: gridlines are usually rendered via CSS borders.
        // If the HTML does not contain "border" style definitions for cells, we assume gridlines are absent.
        bool hasBorderStyles = htmlContent.IndexOf("border", StringComparison.OrdinalIgnoreCase) >= 0;

        if (!hasBorderStyles)
        {
            Console.WriteLine("Gridlines are absent in the output HTML.");
        }
        else
        {
            Console.WriteLine("Gridlines appear to be present in the output HTML.");
        }
    }
}
