// Title: How to move a specific pivot item two positions forward in its row field with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move pivot item two positions forward | reorder row field items in an Excel pivot table using Aspose.Cells | C# code to shift a pivot item ahead in the same field with Aspose.Cells | using PivotItem.Move to change item order in a .xlsx file | programmatically adjust pivot item position in Aspose.Cells workbook
// Tags: Aspose.Cells PivotItem.Move usage | C# change pivot item order programmatically | Excel pivot table item reposition .NET | move pivot item within same field Aspose | adjust pivot item position in workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, locates a pivot item named "ItemName" in the first row field of the first pivot table, moves that item two positions forward within the same parent field using PivotItem.Move, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            PivotTable pivotTable = sheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Pivot table has no row fields.");
                return;
            }

            PivotField rowField = pivotTable.RowFields[0];

            // Identify the pivot item to move
            const string itemName = "ItemName";

            // Find the pivot item by name (manual search because PivotItemCollection does not support LINQ)
            PivotItem item = null;
            foreach (PivotItem pi in rowField.PivotItems)
            {
                if (pi.Name == itemName)
                {
                    item = pi;
                    break;
                }
            }

            if (item == null)
            {
                Console.WriteLine($"Pivot item '{itemName}' not found.");
                return;
            }

            // Current position of the item
            int currentIndex = item.Position;

            // Desired new position (two places forward)
            int newIndex = currentIndex + 2;

            // Clamp the new index within valid range
            if (newIndex >= rowField.PivotItems.Count)
                newIndex = rowField.PivotItems.Count - 1;

            // Move the item within the same parent field (isSameParent = true)
            item.Move(newIndex, true);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
