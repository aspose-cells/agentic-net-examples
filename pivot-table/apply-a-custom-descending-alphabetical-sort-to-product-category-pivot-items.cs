// Title: C# example: Sort Product Category pivot field items in descending alphabetical order using Aspose.Cells for .NET
// Common Searches: how to sort pivot table row items alphabetically descending with Aspose.Cells in C# | Aspose.Cells C# change order of pivot field items programmatically | reorder product category items in an Excel pivot table using Aspose.Cells .NET | set custom position for pivot items in a workbook with Aspose.Cells C# | refresh pivot table after modifying item positions using Aspose.Cells
// Tags: Aspose.Cells pivot item descending sort | C# pivot field position manipulation | Excel pivot table custom order Aspose.Cells | refresh pivot table after sorting Aspose.Cells | product category field sorting C#

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable, PivotField, PivotItem

// The sample loads an existing Excel workbook, locates the first pivot table, finds the "Product Category" field in row or column fields, retrieves its pivot items, sorts them by name in descending alphabetical order, updates each item's Position to reflect the new order, refreshes and recalculates the pivot table, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Assume the pivot table to modify is the first one on the sheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Locate the pivot field named "Product Category"
            PivotField productCategoryField = null;

            // Search in row fields
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name == "Product Category")
                {
                    productCategoryField = field;
                    break;
                }
            }

            // If not found in row fields, search in column fields
            if (productCategoryField == null)
            {
                foreach (PivotField field in pivotTable.ColumnFields)
                {
                    if (field.Name == "Product Category")
                    {
                        productCategoryField = field;
                        break;
                    }
                }
            }

            // Proceed only if the field exists
            if (productCategoryField != null)
            {
                // Retrieve all pivot items for the field
                var items = productCategoryField.PivotItems.Cast<PivotItem>();

                // Sort the item names in descending alphabetical order
                var sortedItems = items.OrderByDescending(item => item.Name).ToList();

                // Reassign positions to reflect the new order
                for (int i = 0; i < sortedItems.Count; i++)
                {
                    sortedItems[i].Position = i;
                }

                // Refresh the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }
            else
            {
                Console.WriteLine("Pivot field 'Product Category' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
