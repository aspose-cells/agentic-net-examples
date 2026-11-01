// Title: Delete all pivot tables with names starting with "Temp_" from every worksheet using Aspose.Cells for .NET (C#)
// Common Searches: c# remove pivot tables that start with Temp_ using Aspose.Cells | how to delete temporary pivot tables from an Excel workbook with Aspose.Cells .NET | programmatically filter and remove pivot tables by name prefix in Aspose.Cells | Aspose.Cells remove pivot tables from all worksheets based on naming convention | delete pivot tables named Temp_* in a workbook using C# Aspose.Cells
// Tags: remove pivot tables by name prefix Aspose.Cells C# | iterate worksheets pivot table collection Aspose.Cells | delete temporary pivot tables Aspose.Cells | Aspose.Cells workbook pivot table cleanup | filter pivot tables using StartsWith in C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, iterates through each worksheet, and removes any pivot tables whose names begin with the "Temp_" prefix before saving the modified workbook to a new file.
class DeleteTempPivotTables
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet
                PivotTableCollection pivots = sheet.PivotTables;

                // Iterate backwards to safely remove items while iterating
                for (int i = pivots.Count - 1; i >= 0; i--)
                {
                    PivotTable pivot = pivots[i];

                    // Remove pivot tables whose names start with "Temp_"
                    if (!string.IsNullOrEmpty(pivot.Name) &&
                        pivot.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                    {
                        pivots.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
