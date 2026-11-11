// Title: How to move a PivotItem up within its parent group using PivotItem.Move in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move pivot row item up one position | PivotItem.Move(-1, true) example for reordering pivot items | Shift a pivot field item upward in the same group with Aspose.Cells | Reorder items in a pivot table row field using Aspose.Cells .NET API | C# code to move a specific pivot item to a higher position
// Tags: PivotItem.Move method Aspose.Cells | reorder pivot table items C# | shift pivot item within same parent group | Aspose.Cells pivot item positioning | C# pivot table row field item order

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, accesses the first pivot table on Sheet1, finds the row field named "Category" and its pivot item "ItemB", then moves that item one position upward within the same parent group using itemToMove.Move(-1, true) and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the pivot table.
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that holds the pivot table.
            Worksheet sheet = workbook.Worksheets["Sheet1"];
            if (sheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Get the first pivot table on the worksheet.
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }
            PivotTable pivotTable = sheet.PivotTables[0];

            // Locate the pivot field named "Category".
            PivotField pivotField = null;
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name == "Category")
                {
                    pivotField = field;
                    break;
                }
            }

            if (pivotField == null)
            {
                Console.WriteLine("Pivot field 'Category' not found.");
                return;
            }

            // Find the pivot item with the name "ItemB".
            PivotItem itemToMove = null;
            foreach (PivotItem item in pivotField.PivotItems)
            {
                if (item.Name == "ItemB")
                {
                    itemToMove = item;
                    break;
                }
            }

            if (itemToMove == null)
            {
                Console.WriteLine("Pivot item 'ItemB' not found.");
                return;
            }

            // Move the item one position up within the same parent group.
            // count = -1 (move up), isSameParent = true (stay in the same group).
            itemToMove.Move(-1, true);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
