// Title: Remove Pivot Tables with "Temp_" Prefix from All Worksheets Using Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel workbook with Aspose.Cells, iterates each worksheet, finds PivotTables whose Name begins with "Temp_", safely removes them from the PivotTableCollection, and saves the cleaned file.
// Keywords: Aspose.Cells | C# | .NET | remove pivot tables | Temp_ prefix | PivotTableCollection | delete temporary pivot tables | Excel automation | programmatic pivot table removal
// Common Searches: Aspose.Cells delete pivot tables Temp_ prefix | C# remove specific pivot tables from Excel workbook | How to programmatically delete pivot tables using Aspose.Cells | Remove temporary pivot tables with .NET | Iterate worksheets and delete pivot tables by name Aspose
// Developer Intent: Delete every PivotTable whose name starts with "Temp_" from all worksheets in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Clean up temporary analysis pivot tables before publishing a final report. | Automate removal of placeholder pivot tables generated during data import. | Prepare a workbook for archiving by stripping development‑only pivot tables.
// AI Prompts: Generate C# code with Aspose.Cells to delete pivot tables that match a name pattern. | Explain how to safely iterate and remove items from a PivotTableCollection in Aspose.Cells. | Show an alternative method to filter and delete pivot tables without backward looping.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemoval
{
    // C# example that loads an Excel workbook with Aspose.Cells, iterates each worksheet, finds PivotTables whose Name begins with "Temp_", safely removes them from the PivotTableCollection, and saves the cleaned file.
    public class RemoveTempPivotTables
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the collection of pivot tables on the current worksheet
                    PivotTableCollection pivots = sheet.PivotTables;

                    // Iterate backwards to safely remove items while looping
                    for (int i = pivots.Count - 1; i >= 0; i--)
                    {
                        PivotTable pt = pivots[i];

                        // Remove pivot tables whose names start with "Temp_"
                        if (!string.IsNullOrEmpty(pt.Name) && pt.Name.StartsWith("Temp_"))
                        {
                            pivots.Remove(pt);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
