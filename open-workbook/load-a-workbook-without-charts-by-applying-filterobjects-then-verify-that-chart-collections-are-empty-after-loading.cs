// Title: Load an Excel workbook, clear all chart objects, and confirm empty chart collections using Aspose.Cells for .NET
// Common Searches: how to clear all charts from an Excel workbook using Aspose.Cells C# | Aspose.Cells verify worksheet chart collection is empty after clearing | load Excel file without chart objects with Aspose.Cells .NET | C# code to remove charts from each worksheet in a workbook | check if any charts remain after loading a workbook with Aspose.Cells
// Tags: clear worksheet charts Aspose.Cells | chart collection emptiness check Aspose.Cells | load workbook without chart objects .NET | remove charts from Excel using Aspose.Cells | chart count validation Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through every worksheet to clear its Charts collection, then validates that each worksheet's chart count is zero, outputting a confirmation when all charts have been successfully removed.
class Program
{
    static void Main()
    {
        // Path to the source workbook (replace with actual file path)
        string sourcePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook normally
            Workbook workbook = new Workbook(sourcePath);

            // Remove all chart objects from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Charts.Clear();
            }

            // Verify that each worksheet's chart collection is empty
            bool allChartsRemoved = true;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count > 0)
                {
                    allChartsRemoved = false;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" still contains {sheet.Charts.Count} chart(s).");
                }
            }

            if (allChartsRemoved)
            {
                Console.WriteLine("All chart collections are empty after loading.");
            }
            else
            {
                Console.WriteLine("Some charts were not removed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
