// Title: Iterate through all worksheets and enable automatic calculation mode in Aspose.Cells for .NET after data import
// Common Searches: Aspose.Cells C# set calculation mode to automatic for each worksheet | Iterate over all worksheets and enable automatic formula recalculation in a .NET workbook | How to turn on automatic calculation after importing data using Aspose.Cells | C# loop through worksheets to configure calculation settings with Aspose.Cells | Enable automatic workbook calculation mode after saving with Aspose.Cells
// Tags: Aspose.Cells enable workbook calculation automatic | C# iterate worksheets set calculation options | formula auto-recalculation Aspose.Cells .NET | configure calculation settings after data import | Aspose.Cells workbook calculation configuration C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, loops through every worksheet to set its CalculationOptions.CalculationMode to Automatic, and then saves the workbook, ensuring that all formulas are recalculated automatically after data import.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // ----- Data import operations would be performed here -----
                // Example: workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir!);
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
