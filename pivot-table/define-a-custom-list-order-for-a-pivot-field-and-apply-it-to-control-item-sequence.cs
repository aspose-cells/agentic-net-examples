// Title: How to define a custom list order for a pivot field and apply it in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a pivot table, enables CustomListSort, and sets the Position property of each PivotItem to a specific order using Aspose.Cells. | Write a script that resets pivot item positions, assigns new indices for a desired sequence, refreshes the pivot table, and saves the workbook with Aspose.Cells. | Provide an example that demonstrates custom sorting of row field values in a pivot table by manipulating PivotItem.Position in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set custom order for pivot table row items | How to change the sequence of pivot field values using Aspose.Cells .NET | Custom list sorting for pivot tables in Aspose.Cells example | Reorder pivot items programmatically with Aspose.Cells C# | Enable CustomListSort and assign positions to PivotItem in Aspose.Cells
// Tags: Aspose.Cells custom list sort pivot field | C# set pivot item position Aspose.Cells | pivot table row field ordering .NET | Aspose.Cells refresh pivot data | custom pivot item sequence Excel

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook with fruit data, adds a pivot table, enables CustomListSort, retrieves the row field items, resets their positions, assigns new Position values to achieve the order Banana, Apple, Orange, Pear, refreshes and calculates the pivot table, and saves the workbook as an XLSX file.
    public class PivotCustomListOrderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                // Columns: Fruit | Quantity
                sheet.Cells["A1"].Value = "Fruit";
                sheet.Cells["B1"].Value = "Quantity";

                sheet.Cells["A2"].Value = "Apple";
                sheet.Cells["B2"].Value = 10;

                sheet.Cells["A3"].Value = "Orange";
                sheet.Cells["B3"].Value = 20;

                sheet.Cells["A4"].Value = "Banana";
                sheet.Cells["B4"].Value = 15;

                sheet.Cells["A5"].Value = "Pear";
                sheet.Cells["B5"].Value = 5;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the row field (Fruit) and the data field (Quantity)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Enable built‑in custom list sorting for the pivot table
                pivotTable.CustomListSort = true;

                // Retrieve the row field that contains the fruit items
                PivotField fruitField = pivotTable.RowFields["Fruit"];
                PivotItemCollection items = fruitField.PivotItems;

                // Reset positions to current indices to avoid conflicts
                for (int i = 0; i < items.Count; i++)
                {
                    items[i].Position = i;
                }

                // Apply the desired custom order
                foreach (PivotItem item in items)
                {
                    switch (item.Name)
                    {
                        case "Banana":
                            item.Position = 0;
                            break;
                        case "Apple":
                            item.Position = 1;
                            break;
                        case "Orange":
                            item.Position = 2;
                            break;
                        case "Pear":
                            item.Position = 3;
                            break;
                    }
                }

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();   // Correct API call
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_CustomListOrder_Demo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotCustomListOrderDemo.Run();
        }
    }
}
