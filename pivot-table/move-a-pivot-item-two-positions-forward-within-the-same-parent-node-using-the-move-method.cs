// Title: How to move a PivotItem two positions forward in a row field using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells move pivot table row item to a new index | Aspose.Cells PivotItem.Move shift item forward by two positions | Programmatically reorder items in a pivot table row field with Aspose.Cells | Change the order of pivot items in an Excel workbook using Aspose.Cells .NET | Move first pivot item two places down in a C# Aspose.Cells example
// Tags: Aspose.Cells PivotItem.Move usage | C# reorder pivot table row items | Aspose.Cells change pivot item order | Excel pivot table item reposition .NET | move pivot item within same parent Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable, PivotField, PivotItem

namespace AsposeCellsExamples
{
    // The sample loads an existing workbook, accesses the first worksheet's first pivot table, selects the first row field's first PivotItem, calculates a new index two positions ahead (clamping to collection bounds), moves the item to that index using PivotItem.Move with the sameParent flag, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Ensure the pivot table has at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("Pivot table contains no row fields.");
                    return;
                }

                // Choose the first row field
                PivotField pivotField = pivotTable.RowFields[0];

                // Ensure the row field has at least one item
                if (pivotField.PivotItems.Count == 0)
                {
                    Console.WriteLine("Row field contains no items.");
                    return;
                }

                // Select the first pivot item to move
                PivotItem pivotItem = pivotField.PivotItems[0];
                int currentPosition = pivotItem.Position;

                // Calculate new position (two places forward)
                int newPosition = currentPosition + 2;

                // Clamp new position within collection bounds
                if (newPosition >= pivotField.PivotItems.Count)
                {
                    newPosition = pivotField.PivotItems.Count - 1;
                }

                // Move the pivot item to the new position (same parent)
                pivotItem.Move(newPosition, true);

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
