// Title: Refresh all pivot tables after updating source worksheet cells with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to refresh pivot tables after changing source worksheet data | programmatically update pivot cache in Aspose.Cells after modifying cell values | Aspose.Cells .NET refresh pivot table calculations after editing worksheet | iterate through workbook worksheets to recalculate pivot tables using Aspose.Cells | example code to call RefreshData and CalculateData on Aspose.Cells pivot tables
// Tags: Aspose.Cells RefreshData API | Aspose.Cells CalculateData usage | C# refresh pivot cache programmatically | update pivot tables after worksheet cell change | iterate workbook worksheets to recalc pivots .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads a workbook, changes a cell in the "Data" sheet, then loops through every worksheet to call RefreshData and CalculateData on each pivot table, ensuring the pivot caches and calculations reflect the updated source data before saving the file.
class PivotTableRefreshExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the source worksheet (named "Data") and modify its data
            Worksheet sourceSheet = workbook.Worksheets["Data"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Worksheet \"Data\" not found in the workbook.");
                return;
            }

            // Example modification: change the value of cell B2
            sourceSheet.Cells["B2"].PutValue(123);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the pivot table's cache with the updated source data
                    pivotTable.RefreshData();

                    // Recalculate the pivot table (necessary for some calculations)
                    pivotTable.CalculateData();
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
