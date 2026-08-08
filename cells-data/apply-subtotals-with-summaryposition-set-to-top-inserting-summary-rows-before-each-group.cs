// Title: Aspose.Cells for .NET: Create a Pivot Table with Top Subtotals (ShowSubtotalAtTop)
// Description: Demonstrates how to build a workbook, populate Category/Product/Sales data, add a pivot table on range A1:C7, place the Category field in the row area, enable automatic subtotals, set ShowSubtotalAtTop = true so summary rows appear before each group, add Sales as a data field, refresh the pivot, and save the file.
// Keywords: Aspose.Cells | C# pivot table | ShowSubtotalAtTop | top subtotals | automatic subtotals | .NET Excel | pivot subtotal position | Excel report generation | category subtotal top
// Common Searches: Aspose.Cells set ShowSubtotalAtTop C# | pivot table subtotals at top Aspose.Cells | how to add top subtotals in Aspose.Cells pivot | C# create pivot table with subtotal before group | Aspose.Cells pivot subtotal position example
// Developer Intent: Create a pivot table and configure its row field to display automatic subtotals before each group (top position).
// Use Cases: Generate a sales summary where each category’s subtotal appears before its product rows for quick insight. | Produce financial worksheets that group data by category and place summary rows at the start of each group. | Build automated Excel reports that need top‑positioned subtotals for hierarchical data analysis.
// AI Prompts: Write C# code using Aspose.Cells to create a pivot table from a range and set ShowSubtotalAtTop = true for the row field. | Explain how to enable automatic subtotals and place them before each group in an Aspose.Cells pivot table. | Provide a step‑by‑step guide to add a pivot table, assign row and data fields, and configure top subtotals with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsSubtotalTopDemo
{
    // Demonstrates how to build a workbook, populate Category/Product/Sales data, add a pivot table on range A1:C7, place the Category field in the row area, enable automatic subtotals, set ShowSubtotalAtTop = true so summary rows appear before each group, add Sales as a data field, refresh the pivot, and save the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Columns: Category, Product, Sales
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                { "Bikes", "Mountain", 1200 },
                { "Bikes", "Road", 1500 },
                { "Cars", "Sedan", 2000 },
                { "Cars", "SUV", 2500 },
                { "Bikes", "Hybrid", 1300 },
                { "Cars", "Coupe", 2200 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    sheet.Cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Add a pivot table based on the data range A1:C7, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add Category as a row field
            int rowFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            PivotField rowField = pivotTable.RowFields[rowFieldPos];

            // Configure the row field to show subtotals at the top (summary rows before each group)
            rowField.ShowSubtotalAtTop = true;   // Insert summary rows before each group
            rowField.IsAutoSubtotals = true;     // Enable automatic subtotals (e.g., Sum)

            // Add Sales as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotSubtotalTopDemo.xlsx");
        }
    }
}
