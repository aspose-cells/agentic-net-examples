// Title: Load an HTML file into an Aspose.Cells workbook without preserving DIV layout using C#
// AI Prompts: Load an HTML document into a Workbook with HtmlLoadOptions while keeping EnableDivTagLayout disabled. | Convert an HTML file to XLSX in C# with Aspose.Cells, ensuring DIV tags are not retained. | Add a pre‑load file‑existence check before importing HTML into an Aspose.Cells workbook.
// Common Searches: C# Aspose.Cells load html file without preserving div tag layout | How to disable EnableDivTagLayout when importing HTML to a workbook in Aspose.Cells | Convert HTML to XLSX using Aspose.Cells while ignoring div elements | Aspose.Cells HtmlLoadOptions default behavior for div layout | Check file existence before loading HTML into Aspose.Cells workbook C#
// Tags: HtmlLoadOptions disable div layout | load html workbook Aspose.Cells C# | html to xlsx conversion Aspose.Cells | file existence check Aspose.Cells | EnableDivTagLayout false Aspose.Cells | Aspose.Cells HTML import settings

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an HTML file into an Aspose.Cells Workbook in C# with HtmlLoadOptions (EnableDivTagLayout left false), includes a file‑existence check, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: Input file not found at '{htmlPath}'.");
            return;
        }

        try
        {
            // Configure load options (default settings do not preserve DIV tag layout)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // (Optional) Save the workbook to another format, e.g., XLSX
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
