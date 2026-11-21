// Title: Disable automatic style creation in Aspose.Cells WorkbookSettings to speed up RemoveUnusedStyles (C#)
// Common Searches: Aspose.Cells C# disable automatic style creation before RemoveUnusedStyles | Set WorkbookSettings.EnableAutomaticStyleCreation false for faster style cleanup | How to prevent new styles from being generated when cleaning up an Excel workbook with Aspose.Cells | Improve performance of RemoveUnusedStyles by turning off automatic style creation in C# | C# Aspose.Cells workbook settings tip for quicker unused style removal
// Tags: WorkbookSettings.EnableAutomaticStyleCreation false | Aspose.Cells RemoveUnusedStyles performance optimization | Excel workbook style cleanup configuration C# | Aspose.Cells workbook settings property usage | C# style management with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRunner
{
    // The example shows how to configure Aspose.Cells WorkbookSettings to set EnableAutomaticStyleCreation to false, then call RemoveUnusedStyles for faster cleanup of unused styles, and finally save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook from the input file
                var workbook = new Workbook(inputPath);

                // Remove all styles that are not used in the workbook
                workbook.RemoveUnusedStyles();

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed successfully. Saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
