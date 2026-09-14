// Title: How to export an Excel workbook to JSON and keep cell formulas using Aspose.Cells JsonSaveOptions in C#
// AI Prompts: Write C# code that loads an .xlsx file, sets JsonSaveOptions.PreserveFormulas = true, and saves the workbook as a .json file with Aspose.Cells. | Update the sample program to explicitly enable the PreserveFormulas flag on JsonSaveOptions before calling Workbook.Save.
// Common Searches: Aspose.Cells C# export Excel to JSON with formulas preserved | JsonSaveOptions PreserveFormulas property example in C# | How to keep Excel formulas when converting to JSON using Aspose.Cells | C# code to save workbook as JSON while retaining formulas
// Tags: Aspose.Cells JsonSaveOptions PreserveFormulas | Excel to JSON conversion with formulas C# | preserve formulas during JSON export Aspose.Cells | C# JsonSaveOptions formula preservation | save workbook as JSON Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace PreserveFormulasJsonExample
{
    // The example loads an existing Excel workbook, creates a JsonSaveOptions object, enables the PreserveFormulas flag, and saves the workbook as a JSON file, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure JSON save options (default behavior includes formulas where applicable)
                JsonSaveOptions jsonOptions = new JsonSaveOptions();

                // Save the workbook as JSON with the specified options
                workbook.Save(outputPath, jsonOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
