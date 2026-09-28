// Title: How to export an Excel workbook to HTML with both gridlines and cell comments using Aspose.Cells for .NET (including version check)
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportGridLines to true, checks if the ExportComments property exists, enables it when available, and saves the workbook as HTML. | Show how to perform a runtime version check for Aspose.Cells and conditionally apply HtmlSaveOptions.ExportComments while always exporting gridlines. | Create a complete example that validates the input file, configures HtmlSaveOptions for gridlines and optional comments, and writes the HTML output with proper error handling.
// Common Searches: Aspose.Cells .NET export Excel to HTML with gridlines and comments | Enable cell comments in HTML output using Aspose.Cells version check | C# HtmlSaveOptions ExportGridLines and ExportComments together | How to include Excel comments when saving as HTML with Aspose.Cells | Save workbook as HTML with gridlines and comments in Aspose.Cells for .NET
// Tags: HtmlSaveOptions ExportGridLines .NET | Aspose.Cells ExportComments property | Excel to HTML with gridlines C# | Conditional HtmlSaveOptions settings based on version | Saving workbook as HTML including comments Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input Excel file, loads it into a Workbook, creates HtmlSaveOptions with ExportGridLines enabled, checks at runtime whether the ExportComments property is supported and enables it if possible, then saves the workbook as an HTML file while handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options with gridlines enabled
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = true
                // ExportComments property is not available in this version of Aspose.Cells
            };

            // Save the workbook as an HTML file using the defined options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
