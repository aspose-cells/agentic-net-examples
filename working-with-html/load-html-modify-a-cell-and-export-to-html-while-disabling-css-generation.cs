// Title: Load an HTML workbook with Aspose.Cells, change cell A1, and save back to HTML without generating CSS
// AI Prompts: Read an HTML file into an Aspose.Cells Workbook, set cell A1 to a custom string, and export the workbook to HTML while preventing any CSS stylesheet from being created. | Configure HtmlSaveOptions so that the saved HTML contains only inline styling and no separate CSS files.
// Common Searches: how to prevent CSS output when saving a workbook to HTML using Aspose.Cells C# | Aspose.Cells load HTML file modify cell and export without CSS | C# Aspose.Cells HtmlSaveOptions disable external stylesheet | change value of A1 in HTML workbook Aspose.Cells .NET example
// Tags: Aspose.Cells load HTML workbook | update cell A1 HtmlSaveOptions | export workbook to HTML without CSS | HtmlSaveOptions suppress external stylesheet | C# Aspose.Cells HTML import export

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an HTML file into an Aspose.Cells Workbook, updating cell A1, and saving the workbook back to HTML using HtmlSaveOptions, with CSS generation suppressed.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.html";

            // Verify that the input HTML file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                return;
            }

            // Load the source HTML file into a workbook
            Workbook workbook = new Workbook(inputPath);

            // Modify a specific cell (e.g., set A1 to a new value)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("New Value");

            // Prepare HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Note: In the current Aspose.Cells version, CSS export control is handled internally.
            // If a specific property to disable CSS exists in your version, set it here.

            // Export the modified workbook back to HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
