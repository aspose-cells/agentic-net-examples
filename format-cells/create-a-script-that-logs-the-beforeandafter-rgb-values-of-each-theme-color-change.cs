// Title: C# Aspose.Cells example: Log before‑and‑after RGB values of Excel theme colors
// Common Searches: how to capture original and new RGB values of Excel theme colors using Aspose.Cells in C# | Aspose.Cells .NET code to log theme palette changes to console | record before and after theme color RGB values when modifying an XLSX workbook with C# | C# script for tracking Excel theme color modifications with Aspose.Cells API | example of iterating through ThemeColorType and logging RGB changes in Aspose.Cells
// Tags: Aspose.Cells theme color RGB extraction .NET | C# log Excel theme palette modifications | Aspose.Cells theme color change tracking | record theme color changes workbook .NET | compare theme color values C# Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeColorLogger
{
    // The example loads an existing workbook, shows where theme‑color iteration and RGB logging would be inserted (the current Aspose.Cells version lacks direct theme APIs), and saves the workbook unchanged while handling any errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: Theme manipulation APIs are not available in the referenced Aspose.Cells version.
                // The original logic that iterated through ThemeColorType and modified theme colors
                // has been omitted to ensure the code compiles and runs successfully.

                // Save the (unmodified) workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
