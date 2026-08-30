// Title: Move a PivotItem up one position within its parent group using PivotItem.Move in Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to shift a specific row field PivotItem upward by one position while keeping it under the same parent node. | Show how to reorder pivot table row items in a .NET workbook by calling PivotItem.Move with a negative offset and the isSameParent flag set to true.
// Common Searches: how to move a pivot table row item up in Aspose.Cells C# | Aspose.Cells PivotItem.Move negative count same parent example | reorder pivot row fields programmatically using Aspose.Cells .NET | C# code to change order of pivot items in Aspose.Cells workbook | move second pivot item upward Aspose.Cells pivot table
// Tags: pivotitem.move upward Aspose.Cells | reorder pivot row field items .NET | aspnet pivot table item hierarchy | c# shift pivot item position | aspose.cells pivot item ordering

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data and a pivot table, retrieves the row field's PivotItem collection, moves the second item up one slot within the same parent using items[1].Move(-1, true), recalculates the pivot, prints the new order, and saves the result as an .xlsx file.
    public class PivotItemMoveUpwardDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["A3"].PutValue("Vegetable");
                sheet.Cells["A4"].PutValue("Fruit");
                sheet.Cells["A5"].PutValue("Vegetable");

                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot cache and calculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Get the PivotItem collection for the row field
                PivotItemCollection items = pivotTable.RowFields["Category"].PivotItems;

                // Display original order
                Console.WriteLine("Original order of pivot items:");
                foreach (PivotItem item in items)
                {
                    Console.WriteLine(item.Name);
                }

                // Move the second item up by one position within the same parent node
                if (items.Count > 1)
                {
                    items[1].Move(-1, true);
                }

                // Recalculate after moving the item
                pivotTable.CalculateData();

                // Display new order
                Console.WriteLine("\nOrder after moving the second item upward:");
                foreach (PivotItem item in items)
                {
                    Console.WriteLine(item.Name);
                }

                // Save the workbook
                string outputPath = "PivotItemMoveUpwardDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Run error: " + ex.Message);
            }
        }
    }
}
