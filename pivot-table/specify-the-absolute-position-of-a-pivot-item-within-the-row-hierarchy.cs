// Title: How to set the absolute position of a pivot item in a row field using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set pivot row item position to first | change order of pivot table row items programmatically with Aspose.Cells | move a specific pivot item to top of row hierarchy in .NET | C# code to reorder pivot field items using Aspose.Cells | set PivotItem.Position property in Aspose.Cells example
// Tags: pivot item position Aspose.Cells C# | row field item reorder Aspose.Cells | set PivotItem.Position .NET | modify pivot table row hierarchy Aspose.Cells | Aspose.Cells move pivot row item | C# Excel pivot table item ordering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, accesses the first pivot table, finds the row field named "Category", locates the pivot item "Beverages", sets its Position property to 0 to place it at the top of the row hierarchy, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one pivot table
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Name of the row field whose item position we want to set
            string rowFieldName = "Category";

            // Locate the row field in the pivot table
            PivotField rowField = null;
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name.Equals(rowFieldName, StringComparison.OrdinalIgnoreCase))
                {
                    rowField = field;
                    break;
                }
            }

            if (rowField != null)
            {
                // Name of the pivot item to reposition
                string itemName = "Beverages";

                // Locate the pivot item within the row field
                PivotItem targetItem = null;
                foreach (PivotItem item in rowField.PivotItems)
                {
                    if (item.Name.Equals(itemName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetItem = item;
                        break;
                    }
                }

                if (targetItem != null)
                {
                    // Move the item to the first position (0‑based)
                    targetItem.Position = 0;
                }
                else
                {
                    Console.WriteLine($"Pivot item '{itemName}' not found in field '{rowFieldName}'.");
                }
            }
            else
            {
                Console.WriteLine($"Row field '{rowFieldName}' not found in the pivot table.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
