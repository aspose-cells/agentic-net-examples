// Title: Remove pivot tables whose names start with "Temp_" from all worksheets in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates every worksheet, and deletes any PivotTable whose Name begins with the prefix "Temp_" before saving the file. | Show how to safely traverse a PivotTableCollection in reverse order to remove matching pivot tables without causing collection modification errors. | Provide a reusable method that accepts input and output file paths and removes temporary pivot tables (prefix "Temp_") from a workbook using Aspose.Cells.
// Common Searches: aspnet remove pivot tables with specific prefix from workbook | c# aspose.cells delete temporary pivot tables across all sheets | how to iterate pivot table collection backwards in Aspose.Cells | programmatically clean up pivot tables named Temp_* in Excel using Aspose.Cells
// Tags: aspose.cells delete pivot tables by name prefix | c# remove temporary pivot tables from Excel workbook | pivot table collection reverse iteration aspose.cells | excel workbook cleanup aspose.cells .net | save workbook after pivot table removal aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemoval
{
    // The example loads InputWorkbook.xlsx, loops through each worksheet, and removes any PivotTable whose Name starts with "Temp_" (case‑insensitive) by iterating the PivotTableCollection in reverse. The modified workbook is then saved as OutputWorkbook.xlsx.
    public class RemoveTempPivotTables
    {
        // Entry point required for console application
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

                    // Loop backwards to safely remove items while iterating
                    for (int i = pivots.Count - 1; i >= 0; i--)
                    {
                        PivotTable pt = pivots[i];

                        // Remove pivot tables whose names start with "Temp_"
                        if (!string.IsNullOrEmpty(pt.Name) &&
                            pt.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                        {
                            pivots.Remove(pt);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
