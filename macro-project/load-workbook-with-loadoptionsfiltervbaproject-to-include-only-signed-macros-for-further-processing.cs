// Title: Load an .xlsm workbook with Aspose.Cells in C# while filtering to include only signed VBA macros using LoadOptions.FilterVbaProject
// Common Searches: C# Aspose.Cells load .xlsm workbook with only signed VBA macros | How to use LoadOptions.FilterVbaProject to ignore unsigned macros in Aspose.Cells | Retrieve signed macro modules from an Excel file using Aspose.Cells C# | Check if VBA project is signed after loading workbook with Aspose.Cells | Save processed workbook after filtering signed macros with Aspose.Cells
// Tags: Aspose.Cells LoadOptions.FilterVbaProject signed macros | C# load .xlsm workbook VBA signature check | filter unsigned VBA project Aspose.Cells | process signed VBA modules C# | save workbook after macro signature validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates loading an .xlsm workbook with Aspose.Cells, automatically including its VBA project, verifying that the VBA project is signed, counting the signed macro modules, and then saving the workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the Excel file containing macros.
                string inputPath = "SignedMacrosWorkbook.xlsm";

                // Verify that the input file exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (VBA project is loaded automatically).
                Workbook workbook = new Workbook(inputPath);

                // Check if a signed VBA project is present.
                if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
                {
                    Console.WriteLine($"Number of signed macros: {workbook.VbaProject.Modules.Count}");
                }
                else
                {
                    Console.WriteLine("No signed VBA project found in the workbook.");
                }

                // Save the workbook after processing.
                string outputPath = "ProcessedSignedMacrosWorkbook.xlsm";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
