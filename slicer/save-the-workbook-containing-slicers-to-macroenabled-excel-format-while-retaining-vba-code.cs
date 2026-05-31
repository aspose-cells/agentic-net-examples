using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    class SaveWorkbookWithMacros
    {
        static void Main()
        {
            try
            {
                // Path to the input workbook that contains slicers and VBA macros
                string inputPath = "input_with_slicers.xlsm";

                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Path for the output macro‑enabled workbook
                string outputPath = "output_with_macros.xlsm";

                // Save the workbook preserving VBA macros
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}