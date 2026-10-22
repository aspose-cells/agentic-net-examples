// Title: How to load an HTML file into an Aspose.Cells workbook in C# without preserving DIV layout
// AI Prompts: Load an HTML document into a Workbook with EnableDivTagLayout set to false using Aspose.Cells. | Convert an HTML file to XLSX in C# while ignoring DIV tags with Aspose.Cells. | Import HTML into a workbook and save as Excel without retaining the original DIV structure.
// Common Searches: Aspose.Cells C# load html without preserving div layout | disable EnableDivTagLayout when importing HTML to a workbook with Aspose.Cells | convert HTML to Excel using Aspose.Cells ignoring div tags
// Tags: Aspose.Cells HtmlLoadOptions disable div layout | C# load html into workbook without div tags | Convert html to xlsx with EnableDivTagLayout false | Aspose.Cells import html ignoring div structure

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions with EnableDivTagLayout disabled (default), and saves the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the HTML file to be loaded.
            string htmlPath = "input.html";

            // Ensure the input file exists.
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load HTML with default options (DIV layout is not preserved by default).
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook.
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook to verify the load.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
