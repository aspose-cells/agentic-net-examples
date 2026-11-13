// Title: Move a pivot table row field and recalculate dependent totals with Aspose.Cells for C#
// Common Searches: C# Aspose.Cells move pivot row field and refresh totals | How to recalculate worksheet formulas after changing pivot layout using Aspose.Cells | Programmatically reposition pivot table items in a .NET workbook | Refresh pivot cache and update dependent formulas with Aspose.Cells | Update Excel pivot row fields and recalc totals in C# code
// Tags: aspocells move pivot row field c# | aspocells refresh pivot cache calculate data | aspocells recalculate worksheet formulas after pivot change | c# reposition pivot table item aspocells | excel dependent totals update aspocells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPivotExample
{
    // The example loads an Excel workbook, verifies a pivot table exists, moves a row field item to a new position, refreshes the pivot cache, recalculates the pivot data, updates any worksheet formulas that depend on the pivot results, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                // Access the first worksheet (adjust index if needed)
                var worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Get the first pivot table
                var pivotTable = worksheet.PivotTables[0];

                // Example: Reposition a row field item.
                // Move the item at index 2 to index 0 in the RowFields collection.
                if (pivotTable.RowFields.Count > 2)
                {
                    // Remove the field from its current position
                    var fieldToMove = pivotTable.RowFields[2];
                    pivotTable.RowFields.RemoveAt(2);

                    // Insert the field at the new position (index 0)
                    pivotTable.RowFields.Insert(0, fieldToMove);
                }

                // Refresh the pivot cache and recalculate the pivot data
                pivotTable.RefreshData();      // Refreshes the source data for the pivot table
                pivotTable.CalculateData();    // Recalculates the pivot table values (including totals)

                // Recalculate any worksheet formulas that depend on the pivot table results
                workbook.CalculateFormula();

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
