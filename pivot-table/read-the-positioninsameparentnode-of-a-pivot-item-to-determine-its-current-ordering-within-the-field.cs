// Title: C# – Read PivotItem.PositionInSameParentNode to Get Item Order in an Aspose.Cells Pivot Table
// Description: Loads a workbook, finds the first pivot table, selects a row field, iterates its PivotItemCollection and prints each item's Name with the PositionInSameParentNode value, which indicates the item's current order within the same parent node.
// Keywords: Aspose.Cells | C# | PivotTable | PivotItem | PositionInSameParentNode | item order | read pivot item index | pivot field position | Aspose.Cells example | Excel automation
// Common Searches: Aspose.Cells get pivot item order C# | PositionInSameParentNode property example | how to read pivot item position with Aspose.Cells | C# retrieve pivot item index Aspose | determine pivot row field item sequence
// Developer Intent: Obtain the PositionInSameParentNode of each PivotItem to identify its current sequence within the selected pivot field.
// Use Cases: Log or display pivot item names alongside their order for debugging. | Compare item positions before and after applying a custom sort. | Programmatically reorder pivot items or apply conditional logic based on their sequence.
// AI Prompts: Write C# code that sorts pivot items by a custom rule and then reads their PositionInSameParentNode using Aspose.Cells. | Show how to move a specific pivot item to a target PositionInSameParentNode in an Aspose.Cells pivot table. | Explain how to synchronize item ordering across multiple pivot tables by comparing PositionInSameParentNode values.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Loads a workbook, finds the first pivot table, selects a row field, iterates its PivotItemCollection and prints each item's Name with the PositionInSameParentNode value, which indicates the item's current order within the same parent node.
    class ReadPivotItemPosition
    {
        static void Main()
        {
            const string inputFile = "PivotData.xlsx";
            const string outputFile = "PivotData_Output.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found. Please ensure the file exists in the application directory.");
                    return;
                }

                // Load the workbook that contains a pivot table
                Workbook workbook = new Workbook(inputFile);

                // Assume the pivot table is on the first worksheet (adjust as needed)
                Worksheet pivotSheet = workbook.Worksheets[0];

                // Get the first pivot table on the sheet
                if (pivotSheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                PivotTable pivotTable = pivotSheet.PivotTables[0];

                // Choose the pivot field whose items' ordering you want to inspect.
                // Here we use the field named "Item". Replace with your actual field name.
                PivotField targetField = pivotTable.RowFields["Item"];
                if (targetField == null)
                {
                    Console.WriteLine("The specified pivot field 'Item' was not found in the row fields.");
                    return;
                }

                // Access the collection of pivot items for the chosen field
                PivotItemCollection items = targetField.PivotItems;

                // Iterate through each pivot item and read its PositionInSameParentNode property
                foreach (PivotItem item in items)
                {
                    int positionInSameParent = item.PositionInSameParentNode;
                    Console.WriteLine($"Pivot Item: {item.Name}, PositionInSameParentNode: {positionInSameParent}");
                }

                // Save the workbook (no changes made to the pivot table in this example)
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
