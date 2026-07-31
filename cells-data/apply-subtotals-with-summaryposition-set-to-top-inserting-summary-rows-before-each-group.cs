// Title: Aspose.Cells C# Pivot Table – Show Subtotals at the Top of Each Category
// Description: Demonstrates how to create a workbook, populate it with sales data, add a pivot table on range A1:C7, place the "Category" field in the row area, enable ShowSubtotalAtTop and automatic subtotals, add the "Sales" field as a data item, refresh the pivot, and save the result as SubtotalAtTopDemo.xlsx.
// Keywords: Aspose.Cells pivot table subtotal top | C# ShowSubtotalAtTop property | Aspose.Cells automatic subtotals | insert summary rows before group | pivot table row subtotals C# | Aspose.Cells .NET example | sales report subtotal at top | global Aspose.Cells tutorial
// Common Searches: Aspose.Cells show subtotal at top C# | pivot table subtotals before each group .NET | how to add top‑position subtotals in Aspose.Cells | C# code for pivot table subtotal rows above categories | Aspose.Cells example for summary rows at top
// Developer Intent: Add a pivot table and configure its row field so that subtotal rows appear before each category group.
// Use Cases: Sales dashboards that need category totals displayed above item details. | Financial statements where section totals precede transaction lines. | Inventory reports that highlight warehouse totals before listing individual SKUs.
// AI Prompts: Generate C# code using Aspose.Cells to create a pivot table with subtotals displayed at the top of each row group. | Explain the effect of ShowSubtotalAtTop and IsAutoSubtotals on pivot table layout in Aspose.Cells. | Provide step‑by‑step instructions to modify an existing Aspose.Cells pivot table so subtotal rows appear before each group.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate it with sales data, add a pivot table on range A1:C7, place the "Category" field in the row area, enable ShowSubtotalAtTop and automatic subtotals, add the "Sales" field as a data item, refresh the pivot, and save the result as SubtotalAtTopDemo.xlsx.
    public class SubtotalAtTopDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                // Header row
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Product");
                worksheet.Cells["C1"].PutValue("Sales");

                // Data rows
                object[,] data = new object[,]
                {
                    { "Electronics", "TV", 1200 },
                    { "Electronics", "Radio", 300 },
                    { "Clothing", "Shirt", 45 },
                    { "Clothing", "Pants", 80 },
                    { "Electronics", "Laptop", 1500 },
                    { "Clothing", "Hat", 20 }
                };

                for (int r = 0; r < data.GetLength(0); r++)
                {
                    for (int c = 0; c < data.GetLength(1); c++)
                    {
                        worksheet.Cells[r + 1, c].PutValue(data[r, c]);
                    }
                }

                // Add a pivot table based on the data range A1:C7, place it at E3
                int pivotIndex = worksheet.PivotTables.Add("A1:C7", "E3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add "Category" as a row field
                int rowFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                PivotField rowField = pivotTable.RowFields[rowFieldPos];

                // Configure the row field to show subtotals at the top (summary rows before each group)
                rowField.ShowSubtotalAtTop = true;   // Insert summary rows before each group
                rowField.IsAutoSubtotals = true;    // Enable automatic subtotals

                // Add "Sales" as a data field (default aggregation is Sum)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("SubtotalAtTopDemo.xlsx");
                Console.WriteLine("Workbook saved as SubtotalAtTopDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SubtotalAtTopDemo.Run();
        }
    }
}
