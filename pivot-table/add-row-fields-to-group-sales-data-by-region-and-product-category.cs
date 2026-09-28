// Title: How to create an Excel pivot table that groups sales by region and product category and sums sales using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that builds a pivot table, adds Region and Category as row fields, and sets Sales as a summed data field. | Show the steps to define a source range, insert a pivot table at a specific cell, then refresh and calculate it using Aspose.Cells. | Provide an example that saves the resulting workbook as an .xlsx file after creating the grouped pivot.
// Common Searches: aspocells c# create pivot table with multiple row fields | group sales data by region and category in Excel using Aspose.Cells | sum sales column in Aspose.Cells pivot table example | how to refresh and calculate a pivot table with Aspose.Cells .NET | save pivot table workbook as xlsx using Aspose.Cells C#
// Tags: Aspose.Cells pivot table row grouping | C# define source range for Aspose.Cells pivot | Aspose.Cells sum function for data field | export Aspose.Cells workbook to XLSX | refresh and calculate Aspose.Cells pivot programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // The sample creates a new workbook, fills it with sales data (Region, Category, Sales), defines the source range A1:C9, adds a pivot table at E3 named SalesPivot, adds Region and Category as row fields, adds Sales as a data field with SUM aggregation, refreshes and calculates the pivot, and saves the file as GroupedSalesPivot.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data
            // Header row
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Sales");

            // Data rows
            object[,] data = new object[,]
            {
                { "North", "Electronics", 1200 },
                { "North", "Furniture",   800 },
                { "South", "Electronics", 1500 },
                { "South", "Furniture",   700 },
                { "East",  "Electronics", 900 },
                { "East",  "Furniture",   600 },
                { "West",  "Electronics", 1100 },
                { "West",  "Furniture",   500 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    sheet.Cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Define the source range for the pivot table
            string sourceRange = "A1:C9";

            // Add a pivot table to the worksheet at cell E3
            int pivotIndex = sheet.PivotTables.Add(sourceRange, "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add row fields to group by Region and then by Category
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the Sales field as a data field (sum aggregation)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table to populate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the pivot table
            workbook.Save("GroupedSalesPivot.xlsx");
        }
    }
}
