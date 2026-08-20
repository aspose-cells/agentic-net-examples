// Title: Replace an Existing Pivot Table with a New One Using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, removes the first pivot table on the first worksheet (if present), creates a new pivot at D5 using the used range, sets the first column as rows and the second as values, refreshes and calculates the pivot, then saves the file.
// Keywords: Aspose.Cells C# remove pivot | Aspose.Cells add pivot table | replace pivot table .NET | delete and recreate pivot Aspose | refresh pivot data programmatically | Excel pivot table automation C#
// Common Searches: Aspose.Cells delete pivot table C# | How to replace a pivot table with Aspose.Cells | Create new pivot table after removing old one .NET | Refresh pivot cache after replacement Aspose | Programmatic pivot table replacement in Excel using C#
// Developer Intent: Remove a pivot table from a worksheet and insert a newly configured pivot table in its place using Aspose.Cells for .NET.
// Use Cases: Automate monthly reports by deleting the previous pivot and rebuilding it with the latest data range. | Provide a template workbook that always generates a fresh pivot table, regardless of prior content. | Process multiple workbooks in a batch, standardizing pivot layouts by removing existing pivots and adding a uniform new one.
// AI Prompts: Write C# code with Aspose.Cells that deletes the first pivot table on a sheet and adds a new pivot at D5, using column 0 as rows and column 1 as values. | Show how to loop through all worksheets in a workbook, remove any existing pivots, create a new pivot based on the used range, refresh it, and save the file. | Explain how to safely handle workbooks that contain no pivot tables before attempting removal, then add a new pivot table programmatically.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotReplaceDemo
{
    // Loads an Excel workbook, removes the first pivot table on the first worksheet (if present), creates a new pivot at D5 using the used range, sets the first column as rows and the second as values, refreshes and calculates the pivot, then saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWithPivot.xlsx";
                const string outputPath = "OutputWithNewPivot.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains the original pivot table
                Workbook workbook = new Workbook(inputPath);

                // Work with the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Remove the existing pivot table (if any)
                // -------------------------------------------------
                PivotTableCollection pivots = sheet.PivotTables;
                if (pivots.Count > 0)
                {
                    // Remove the first pivot table and its data
                    PivotTable existingPivot = pivots[0];
                    pivots.Remove(existingPivot);
                }

                // -------------------------------------------------
                // 2. Add a new pivot table with the same source data
                // -------------------------------------------------
                // Determine the source data range (used range of the worksheet)
                Aspose.Cells.Range sourceRange = sheet.Cells.MaxDisplayRange;
                string sourceData = $"={sheet.Name}!{sourceRange.Address}";

                // Add the new pivot table at cell D5 with a new name
                int newIndex = pivots.Add(sourceData, "D5", "NewPivotTable");
                PivotTable newPivot = pivots[newIndex];

                // -------------------------------------------------
                // 3. Configure the new pivot table (example: Row = first column, Data = second column)
                // -------------------------------------------------
                // Assuming the first column contains categories and the second column contains values
                newPivot.AddFieldToArea(PivotFieldType.Row, 0);   // Row field from first column
                newPivot.AddFieldToArea(PivotFieldType.Data, 1);  // Data field from second column

                // Refresh and calculate to populate the pivot table
                newPivot.RefreshData();          // Refresh source data for the pivot cache
                newPivot.CalculateData();        // Recalculate pivot data

                // -------------------------------------------------
                // 4. Save the workbook with the updated pivot table
                // -------------------------------------------------
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
