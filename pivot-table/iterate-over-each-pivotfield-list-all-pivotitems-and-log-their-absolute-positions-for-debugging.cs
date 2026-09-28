// Title: C# example: enumerate all PivotFields in an Aspose.Cells PivotTable and log each PivotItem’s absolute Position for debugging
// AI Prompts: Write C# code with Aspose.Cells that iterates over every PivotField in a PivotTable, checks for PivotItems, and prints each item's Name and Position to the console. | Demonstrate how to output row, column, data, and page field items together with their absolute positions from an Aspose.Cells PivotTable to aid troubleshooting.
// Common Searches: Aspose.Cells C# get absolute position of pivot items in a PivotTable | How to list all pivot field items with positions using Aspose.Cells .NET | Debugging Aspose.Cells pivot tables by printing row and column field items
// Tags: Aspose.Cells enumerate pivot fields C# | Aspose.Cells pivot item position debugging | C# Aspose.Cells pivot table item iteration | Aspose.Cells log pivot items absolute position

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDebug
{
    // The sample creates a workbook, adds data, builds a PivotTable, then iterates through each PivotField collection (row, column, data, page). For every field that contains PivotItems it prints the item name and its absolute Position, helping developers debug pivot structures, and finally saves the workbook as PivotFieldItemsPositionsDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("A");
                worksheet.Cells["B4"].PutValue(30);

                // Add a pivot table to the worksheet
                int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Helper method to iterate over a collection of PivotFields
                void LogPivotFields(PivotFieldCollection fields)
                {
                    foreach (PivotField field in fields)
                    {
                        Console.WriteLine($"PivotField: {field.Name}");
                        // Some field types (e.g., Data fields) may not have PivotItems
                        if (field.PivotItems != null)
                        {
                            foreach (PivotItem item in field.PivotItems)
                            {
                                // Position property gives the absolute position of the item
                                Console.WriteLine($"  Item: {item.Name}, Position: {item.Position}");
                            }
                        }
                    }
                }

                // Log all pivot fields in each area
                Console.WriteLine("=== Row Fields ===");
                LogPivotFields(pivotTable.RowFields);

                Console.WriteLine("=== Column Fields ===");
                LogPivotFields(pivotTable.ColumnFields);

                Console.WriteLine("=== Data Fields ===");
                LogPivotFields(pivotTable.DataFields);

                Console.WriteLine("=== Page Fields ===");
                LogPivotFields(pivotTable.PageFields);

                // Save the workbook (debug workbook can be inspected if needed)
                workbook.Save("PivotFieldItemsPositionsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
