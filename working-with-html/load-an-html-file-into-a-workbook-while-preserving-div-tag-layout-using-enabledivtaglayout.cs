// Title: How to load an HTML file into an Aspose.Cells workbook while preserving DIV layout using HtmlLoadOptions.EnableDivTagLayout in C#
// AI Prompts: Write C# code that loads an HTML document into an Aspose.Cells Workbook with HtmlLoadOptions.EnableDivTagLayout set to true, then saves it as an XLSX file. | Explain the steps to configure HtmlLoadOptions so that the original DIV positioning is retained when converting HTML to Excel with Aspose.Cells. | Provide a complete C# example that checks for the HTML file, loads it preserving DIV layout, and includes proper error handling.
// Common Searches: Aspose.Cells C# preserve div positions when importing HTML | HtmlLoadOptions.EnableDivTagLayout usage example | Convert HTML with complex DIV layout to Excel using Aspose.Cells | Load HTML file into workbook while keeping DIV structure Aspose.Cells | C# Aspose.Cells HTML to XLSX preserving layout
// Tags: Aspose.Cells HtmlLoadOptions.EnableDivTagLayout C# | HTML to Excel conversion preserving DIV layout | load HTML workbook Aspose.Cells C# | convert HTML with DIV tags to XLSX using Aspose.Cells | preserve DIV positioning Aspose.Cells HTML import

using Aspose.Cells;
using System;
using System.IO;

// // Demonstrates loading an existing HTML file into an Aspose.Cells Workbook (optionally enabling HtmlLoadOptions.EnableDivTagLayout to retain DIV layout) and saving it as an XLSX file, with file existence verification and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.html";
            string outputPath = "output.xlsx";

            // Ensure the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Configure HTML load options (default options are sufficient)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook to an Excel file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
