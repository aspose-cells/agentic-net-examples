// Title: Reorder specific row items in an Excel pivot table with Aspose.Cells for .NET and save the workbook
// AI Prompts: Update the positions of the row items 'Apple', 'Banana', and 'Orange' in a pivot table, recalculate the pivot, and write the result to a new XLSX file using Aspose.Cells C#. | Generate C# code that safely changes the order of pivot items, skips items that are not present, refreshes the pivot data, and saves the workbook as OutputPivot.xlsx.
// Common Searches: Aspose.Cells C# set pivot row item order programmatically | How to change the position of specific items in an Excel pivot table using .NET | Reorder pivot table row fields and save workbook with Aspose.Cells | Handle missing pivot items when updating order in C# | Refresh and recalculate pivot table after modifying item positions Aspose.Cells
// Tags: pivot item position Aspose.Cells | reorder pivot row items C# | save workbook after pivot changes Aspose.Cells | missing pivot item handling .NET | recalculate pivot table Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotExample
{
    // The example loads InputPivot.xlsx, accesses the first worksheet's first pivot table, refreshes its data, safely sets the positions of the row items "Apple", "Banana", and "Orange", refreshes and recalculates the pivot again, and saves the modified workbook as OutputPivot.xlsx.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputPivot.xlsx";
            const string outputPath = "OutputPivot.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Refresh data using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure the pivot table has at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("Pivot table has no row fields.");
                    return;
                }

                // Access the first row field
                PivotField rowField = pivotTable.RowFields[0];

                // Set positions for specific pivot items, handling missing items gracefully
                SetPivotItemPosition(rowField, "Apple", 0);
                SetPivotItemPosition(rowField, "Banana", 1);
                SetPivotItemPosition(rowField, "Orange", 2);

                // Recalculate after modifying item positions
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to set the position of a pivot item safely
        private static void SetPivotItemPosition(PivotField field, string itemName, int position)
        {
            try
            {
                // Find the pivot item by name
                PivotItem pivotItem = field.PivotItems.FirstOrDefault(p => p.Name.Equals(itemName, StringComparison.OrdinalIgnoreCase));

                if (pivotItem != null)
                {
                    pivotItem.Position = position;
                }
                else
                {
                    Console.WriteLine($"Pivot item '{itemName}' not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to set position for '{itemName}': {ex.Message}");
            }
        }
    }
}
