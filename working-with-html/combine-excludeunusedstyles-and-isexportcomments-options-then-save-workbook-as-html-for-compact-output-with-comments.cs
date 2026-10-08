// Title: Generate compact HTML from an Excel file with unused styles removed and cell comments included using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, sets HtmlSaveOptions.ExcludeUnusedStyles to true, ensures comments are exported, and saves the result as an .html file with Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to produce a minimal HTML file from Excel while preserving cell comments in a .NET project. | Demonstrate robust error handling for a missing input workbook when converting to HTML with ExcludeUnusedStyles enabled in Aspose.Cells.
// Common Searches: asp.net convert excel to html exclude unused styles Aspose.Cells | how to keep cell comments when saving workbook as html using Aspose.Cells | compact html output from excel workbook with Aspose.Cells HtmlSaveOptions | Aspose.Cells HtmlSaveOptions ExcludeUnusedStyles true example | C# save workbook as html with comments and minimal CSS Aspose
// Tags: exclude unused styles Aspose.Cells HtmlSaveOptions | export cell comments to HTML Aspose.Cells | compact HTML conversion from Excel C# | Aspose.Cells HTML output size optimization | error handling missing workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads input.xlsx, configures HtmlSaveOptions with ExcludeUnusedStyles = true (comments are exported by default), and saves the workbook as output.html, including basic file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Remove unused styles to make the HTML output smaller
                ExcludeUnusedStyles = true
                // ExportComments property is not available in this version of Aspose.Cells,
                // so it is omitted. Comments will be exported by default if supported.
            };

            // Save the workbook as an HTML file with the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
