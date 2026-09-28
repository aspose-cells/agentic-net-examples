// Title: Load a legacy Excel 97‑2003 (.xls) workbook with Aspose.Cells using LoadOptions and save it in C#
// AI Prompts: Write C# code that verifies an input .xls file exists, opens it with Aspose.Cells using LoadOptions set to LoadFormat.Excel97To2003, and writes the workbook to a specified output path. | Create a reusable method that accepts input and output file names, checks the input file, loads the legacy Excel workbook via LoadOptions.Version = Excel97To2003, and saves the result while handling possible exceptions.
// Common Searches: asp.net core open legacy Excel workbook using Aspose.Cells Excel97To2003 format | c# validate file existence before loading .xls with Aspose.Cells LoadOptions | how to save a workbook after loading it with LoadFormat.Excel97To2003 in C# | using LoadOptions to read Excel 97-2003 files in Aspose.Cells C# example
// Tags: Aspose.Cells LoadOptions Excel97To2003 C# example | load legacy .xls workbook Aspose.Cells | file existence check before Aspose.Cells load | save workbook after LoadOptions load Aspose.Cells | exception handling Aspose.Cells workbook load

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates loading a legacy Excel 97‑2003 (.xls) file with Aspose.Cells by configuring LoadOptions to LoadFormat.Excel97To2003, includes a pre‑load file‑existence check, optional processing placeholder, and saving the workbook to a new .xls file with basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Configure LoadOptions for Excel 97‑2003 format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any required operations on the workbook here

            // Save the workbook to the desired output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
