// Title: Set absolute position of a PivotItem in a row hierarchy with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, accesses the first pivot table, selects a row field, finds a pivot item by name, assigns its PositionInSameParentNode to place it at a specific index, recalculates the pivot table, and saves the updated file.
// Keywords: Aspose.Cells PivotItem PositionInSameParentNode | C# reorder pivot row items | move pivot item to top Aspose | set pivot item absolute position .NET | pivot table item ordering Aspose.Cells
// Common Searches: Aspose.Cells change pivot item order C# | PositionInSameParentNode example | how to move pivot row item to first position | set absolute position of pivot item Aspose | reorder row field items in pivot table .NET
// Developer Intent: Place a specific pivot item at a defined index within the row hierarchy of an Aspose.Cells pivot table.
// Use Cases: Show a priority product category first in a sales dashboard. | Display the current month before other months in a financial report. | Ensure a custom label appears at the top of a hierarchical list before exporting.
// AI Prompts: Generate C# code that sets PositionInSameParentNode for a PivotItem to a given index using Aspose.Cells. | Provide an example that checks for a pivot item’s existence, moves it to the first position, and handles missing items gracefully. | Explain why recalculating the pivot table is required after changing item positions in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    // Loads a workbook, accesses the first pivot table, selects a row field, finds a pivot item by name, assigns its PositionInSameParentNode to place it at a specific index, recalculates the pivot table, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourceFile = "PivotSource.xlsx";
                const string outputFile = "PivotSource_WithItemPosition.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Error: File '{sourceFile}' not found.");
                    return;
                }

                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(sourceFile);

                // Assume the pivot table is on the first worksheet
                Worksheet pivotSheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (pivotSheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = pivotSheet.PivotTables[0];

                // Ensure there is at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("Error: Pivot table does not contain any row fields.");
                    return;
                }

                // Choose the first row field (adjust index as needed)
                PivotField rowField = pivotTable.RowFields[0];

                // Access the collection of pivot items for that row field
                PivotItemCollection items = rowField.PivotItems;

                // Name of the pivot item to reposition
                string targetItemName = "ItemName";

                // Find the pivot item by name
                PivotItem targetItem = null;
                foreach (PivotItem pi in items)
                {
                    if (pi.Name.Equals(targetItemName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetItem = pi;
                        break;
                    }
                }

                if (targetItem != null)
                {
                    // Set the item's position among its siblings (0 = first)
                    targetItem.PositionInSameParentNode = 0;
                }
                else
                {
                    Console.WriteLine($"Warning: Pivot item '{targetItemName}' not found.");
                }

                // Recalculate the pivot table after changing positions
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
