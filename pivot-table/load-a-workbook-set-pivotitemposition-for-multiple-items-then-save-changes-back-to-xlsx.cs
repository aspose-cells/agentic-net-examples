using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDemo
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the first worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Refresh data before making changes
                pivotTable.RefreshData();

                // Access the first row field (ensure at least one row field exists)
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("The pivot table has no row fields.");
                    return;
                }

                PivotField rowField = pivotTable.RowFields[0];
                PivotItemCollection items = rowField.PivotItems;

                // Set positions for specific items
                SetItemPosition(items, "Apple", 0);
                SetItemPosition(items, "Banana", 1);
                SetItemPosition(items, "Orange", 2);

                // Recalculate after modifications
                pivotTable.CalculateData();

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to safely set the Position of a PivotItem if it exists
        private static void SetItemPosition(PivotItemCollection items, string itemName, int newPosition)
        {
            try
            {
                PivotItem item = items[itemName];
                if (item != null)
                {
                    item.Position = newPosition;
                }
            }
            catch
            {
                // Item with the specified name does not exist; ignore
            }
        }
    }
}