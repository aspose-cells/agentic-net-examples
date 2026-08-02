// Title: Read PivotItem.PositionInSameParentNode to Determine Item Order in an Aspose.Cells PivotTable (C#)
// Description: Loads a workbook, accesses the first pivot table on the "PivotTable" sheet, retrieves the row field "Item", iterates its PivotItems, and reads each item's PositionInSameParentNode property to reveal the current ordering within the parent node. The example prints the name and position and saves the workbook.
// Keywords: Aspose.Cells read pivot item position | PivotItem PositionInSameParentNode C# | Aspose.Cells pivot item order | C# Aspose.Cells get item sequence | pivot table item ranking Aspose | Aspose.Cells .NET pivot item position
// Common Searches: How to get pivot item order with Aspose.Cells for .NET | Read PositionInSameParentNode of a PivotItem in C# | Determine row field item sequence in Aspose.Cells | Retrieve pivot item positions from existing workbook Aspose.Cells | Aspose.Cells C# pivot item sorting example
// Developer Intent: Obtain the PositionInSameParentNode value for each PivotItem to identify its current placement within the row field hierarchy.
// Use Cases: Display or log the current ordering of row‑field items for debugging or audit trails. | Compare the existing item order with a custom sort and programmatically rearrange items if needed. | Export a mapping of item names to their positions for integration with external reporting tools.
// AI Prompts: Generate C# code that reorders pivot items in an Aspose.Cells PivotTable based on their PositionInSameParentNode values. | Create a method that returns a dictionary of PivotItem names and their PositionInSameParentNode for a specified row field. | Write a script that logs each PivotItem's PositionInSameParentNode and then moves items to match a predefined sequence.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads a workbook, accesses the first pivot table on the "PivotTable" sheet, retrieves the row field "Item", iterates its PivotItems, and reads each item's PositionInSameParentNode property to reveal the current ordering within the parent node. The example prints the name and position and saves the workbook.
    public class ReadPivotItemPositionInSameParentNode
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputFile = "PivotTest3.xlsx";
            const string outputFile = "PivotItemPositionsRead.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputFile);

            // Try to get the worksheet named "PivotTable"
            Worksheet pivotSheet = workbook.Worksheets["PivotTable"];
            if (pivotSheet == null)
            {
                // If not found, use the first worksheet as a fallback
                pivotSheet = workbook.Worksheets[0];
                Console.WriteLine("Worksheet \"PivotTable\" not found. Using the first worksheet instead.");
            }

            // Ensure there is at least one pivot table on the sheet
            if (pivotSheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the selected worksheet.");
                return;
            }

            // Get the first pivot table (adjust index if needed)
            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // Access the row field named "Item"
            PivotField itemField = pivotTable.RowFields["Item"];
            if (itemField == null)
            {
                Console.WriteLine("Row field \"Item\" not found in the pivot table.");
                return;
            }

            // Iterate through pivot items and display their PositionInSameParentNode
            Console.WriteLine("Pivot Item positions within the same parent node:");
            foreach (PivotItem item in itemField.PivotItems)
            {
                int position = item.PositionInSameParentNode;
                Console.WriteLine($"Item Name: {item.Name}, PositionInSameParentNode: {position}");
            }

            // Save the workbook (optional)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as \"{outputFile}\".");
        }
    }
}
