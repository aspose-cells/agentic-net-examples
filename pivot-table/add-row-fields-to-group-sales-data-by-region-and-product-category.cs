// Title: C# – Add Region and Category Row Fields to an Aspose.Cells Pivot Table and Sum Sales
// Description: Creates a new workbook, populates it with Region, Category, and Sales data, builds a pivot table on A1:C7 at E3, adds Region and Category as row fields, sets Sales as a summed data field, refreshes and calculates the pivot, and saves the result as GroupedSalesPivot.xlsx.
// Keywords: Aspose.Cells C# pivot table example | add row fields Aspose.Cells | group sales by region and category | pivot table sum aggregation C# | Aspose.Cells refresh calculate pivot | Excel pivot table Aspose.Cells .NET | regional sales summary Aspose.Cells
// Common Searches: How to add multiple row fields in an Aspose.Cells pivot table C# | Group sales data by region and category using Aspose.Cells | Set sum function for a data field in Aspose.Cells pivot table | Refresh and calculate pivot table after adding fields Aspose.Cells | Aspose.Cells example for creating pivot tables in .NET
// Developer Intent: Generate a pivot table that first groups sales records by geographic region, then by product category, and shows the total sales for each combination.
// Use Cases: Produce a regional sales summary report with category breakdowns. | Create a dashboard worksheet that visualizes product performance per market. | Export a ready‑to‑share Excel file containing grouped sales totals for stakeholders.
// AI Prompts: Write C# code using Aspose.Cells to add Region and Category as row fields in a pivot table and sum the Sales field. | Show how to change the Sales data field aggregation from Sum to Average in an Aspose.Cells pivot table. | Provide steps to format pivot table headers and data cells after creating the pivot with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // Creates a new workbook, populates it with Region, Category, and Sales data, builds a pivot table on A1:C7 at E3, adds Region and Category as row fields, sets Sales as a summed data field, refreshes and calculates the pivot, and saves the result as GroupedSalesPivot.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data: Region, Category, Sales
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Sales");

            // Sample rows
            var data = new object[,]
            {
                { "North", "Widgets", 1200 },
                { "North", "Gadgets", 800 },
                { "South", "Widgets", 1500 },
                { "South", "Gadgets", 1100 },
                { "East",  "Widgets", 900 },
                { "East",  "Gadgets", 700 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                sheet.Cells[r + 1, 0].PutValue(data[r, 0]); // Region
                sheet.Cells[r + 1, 1].PutValue(data[r, 1]); // Category
                sheet.Cells[r + 1, 2].PutValue(data[r, 2]); // Sales
            }

            // Create a pivot table based on the data range A1:C7, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add row fields to group by Region and then by Category
            pivot.AddFieldToArea(PivotFieldType.Row, "Region");
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the Sales field as a data field and set aggregation to Sum
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.DataFields[0].Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table to populate data
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook with the pivot table
            workbook.Save("GroupedSalesPivot.xlsx");
        }
    }
}
