// Title: How to disable automatic style creation with WorkbookSettings before calling RemoveUnusedStyles in Aspose.Cells for .NET
// Common Searches: disable automatic style creation in Aspose.Cells C# before removing unused styles | configure WorkbookSettings.AutoStyleCreation false to speed up RemoveUnusedStyles | optimize style cleanup performance in Aspose.Cells .NET workbook | Aspose.Cells remove unused styles without generating new styles | C# example for setting WorkbookSettings to prevent new styles during cleanup
// Tags: WorkbookSettings.AutoStyleCreation false | RemoveUnusedStyles performance optimization | Aspose.Cells style cleanup configuration | disable style generation Excel workbook C# | Aspose.Cells workbook settings API

using System;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates setting WorkbookSettings.AutoStyleCreation to false to stop automatic style creation, then invoking RemoveUnusedStyles for fast cleanup of unused styles, and finally saving the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Remove all unused styles from the workbook
            workbook.RemoveUnusedStyles();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
