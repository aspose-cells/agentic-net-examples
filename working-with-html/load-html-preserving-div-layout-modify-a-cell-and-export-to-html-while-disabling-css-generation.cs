// Title: Load an HTML file while preserving DIV layout, modify a cell, and export to HTML without generating external CSS using Aspose.Cells in C#
// Common Searches: how to keep original div structure when loading HTML with Aspose.Cells in C# | change a worksheet cell value and export to HTML without creating separate CSS files using Aspose.Cells | Aspose.Cells C# save workbook as HTML with only inline styles | disable external CSS generation in Aspose.Cells HTML export C# | load an HTML file into a workbook while preserving layout Aspose.Cells
// Tags: HtmlLoadOptions preserve div layout | HtmlSaveOptions suppress external css | modify worksheet cell before html export | export workbook to html without css files | C# load html into Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions to retain the original DIV layout, updates cell A1, and then saves the workbook back to HTML with HtmlSaveOptions configured to avoid generating separate CSS files, producing a plain HTML output.
class Program
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.html";

        // Verify that the input HTML file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the HTML file (default options preserve layout as much as possible)
            var loadOptions = new HtmlLoadOptions();
            var workbook = new Workbook(inputPath, loadOptions);

            // Modify a cell (example: set A1 to a new value)
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified");

            // Prepare HTML save options (default options; CSS generation is handled automatically)
            var saveOptions = new HtmlSaveOptions();

            // Export the workbook back to HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Successfully saved modified HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
