// Title: Aspose.Cells for .NET: Create a Pivot Table that Groups Sales by Region and Product Category (C#)
// Description: C# code that builds a workbook with sample sales rows (Region, Category, Sales), adds a pivot table on A1:C7, assigns Region and Category to the row area, sums the Sales field, refreshes the cache, calculates the results, and saves the file as GroupedSalesByRegionAndCategory.xlsx.
// Keywords: Aspose.Cells | C# pivot table example | add row fields to pivot | group sales by region | product category pivot | sum aggregation Aspose.Cells | Excel automation .NET | pivot cache refresh | calculate pivot data | sample workbook code
// Common Searches: Aspose.Cells add multiple row fields pivot C# | C# pivot table group by region and category | How to sum sales in an Aspose.Cells pivot | Create pivot table programmatically Aspose.Cells .NET
// Developer Intent: Add Region and Category as row fields in a pivot table and compute total Sales using Aspose.Cells for .NET.
// Use Cases: Produce a regional sales summary that breaks down totals by product category for management reporting. | Build an Excel workbook that can be refreshed with new sales data while automatically updating grouped totals. | Export a pivot‑driven sales analysis to Excel for distribution to finance or marketing teams.
// AI Prompts: Show how to add a Year column field to the same pivot table with Aspose.Cells. | Provide C# code to format pivot headers and apply currency formatting to the Sales sum. | Explain how to change the data source range of the pivot table and refresh it programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# code that builds a workbook with sample sales rows (Region, Category, Sales), adds a pivot table on A1:C7, assigns Region and Category to the row area, sums the Sales field, refreshes the cache, calculates the results, and saves the file as GroupedSalesByRegionAndCategory.xlsx.
    public class GroupSalesByRegionAndCategory
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("North");
            sheet.Cells["B3"].PutValue("Furniture");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("South");
            sheet.Cells["B4"].PutValue("Electronics");
            sheet.Cells["C4"].PutValue(1500);

            sheet.Cells["A5"].PutValue("South");
            sheet.Cells["B5"].PutValue("Furniture");
            sheet.Cells["C5"].PutValue(700);

            sheet.Cells["A6"].PutValue("East");
            sheet.Cells["B6"].PutValue("Electronics");
            sheet.Cells["C6"].PutValue(900);

            sheet.Cells["A7"].PutValue("East");
            sheet.Cells["B7"].PutValue("Furniture");
            sheet.Cells["C7"].PutValue(600);

            // Add a pivot table based on the data range A1:C7
            int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure row fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add data field and set aggregation to Sum
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

            // Refresh the pivot cache and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("GroupedSalesByRegionAndCategory.xlsx");
        }
    }
}
