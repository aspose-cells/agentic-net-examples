// Title: Setting the absolute position of a specific pivot item in a row field with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move pivot row field item to first position | Set pivot item Position property in Aspose.Cells .NET example | Change order of a pivot item programmatically using Aspose.Cells | Reorder row hierarchy items in an Excel pivot table with C# | Adjust pivot table row field item index using Aspose.Cells API
// Tags: Aspose.Cells pivot item Position property | C# reorder pivot row field items | Aspose.Cells modify pivot hierarchy order | Excel pivot table item positioning .NET | Aspose.Cells set pivot item index

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, accesses the first worksheet and its first pivot table, finds the row field named "Category" and the pivot item "Beverages", sets the item's zero‑based Position to 0 to move it to the top of the row hierarchy, and saves the workbook, with error handling for missing files and loading issues.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (if the file is password‑protected, an exception will be thrown)
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        try
        {
            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the first pivot table on the worksheet
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Name of the row field that contains the pivot item
            string rowFieldName = "Category";

            // Locate the row field within the pivot table
            PivotField rowField = null;
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name == rowFieldName)
                {
                    rowField = field;
                    break;
                }
            }

            if (rowField != null)
            {
                // Name of the pivot item whose absolute position we want to set
                string pivotItemName = "Beverages";

                // Locate the pivot item within the row field
                PivotItem pivotItem = null;
                foreach (PivotItem item in rowField.PivotItems)
                {
                    if (item.Name == pivotItemName)
                    {
                        pivotItem = item;
                        break;
                    }
                }

                if (pivotItem != null)
                {
                    // Set the absolute (zero‑based) position of the pivot item in the row hierarchy
                    // For example, move it to the first position
                    pivotItem.Position = 0;
                }
                else
                {
                    Console.WriteLine($"Pivot item \"{pivotItemName}\" not found in field \"{rowFieldName}\".");
                }
            }
            else
            {
                Console.WriteLine($"Row field \"{rowFieldName}\" not found in the pivot table.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
