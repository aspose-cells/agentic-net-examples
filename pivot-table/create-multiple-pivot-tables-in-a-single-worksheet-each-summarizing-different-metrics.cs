// Title: Create Multiple Pivot Tables in One Worksheet with Aspose.Cells for .NET
// Description: This example builds a workbook, fills it with random sales data, defines the entire data block as the source range, and adds two pivot tables to the same sheet—one summarizing total Quantity by Region at F3 and another summarizing total Revenue by Product at F20. Each pivot uses a different built‑in style, is refreshed, calculated, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells multiple pivot tables | C# create pivot tables same worksheet | Aspose.Cells .NET pivot table example | summarize quantity by region Aspose | summarize revenue by product Aspose | pivot table style Aspose.Cells | refresh calculate pivot Aspose
// Common Searches: how to add two pivot tables on one sheet using Aspose.Cells | Aspose.Cells example multiple pivot tables .NET | set different styles for each pivot table Aspose | create pivot tables from same source range Aspose.Cells
// Developer Intent: Generate two distinct pivot tables on a single worksheet, each summarizing a different metric from the same source data.
// Use Cases: Generate a sales report that shows quantity per region alongside revenue per product in one workbook. | Build a dashboard worksheet with side‑by‑side KPIs for quick comparative analysis. | Apply separate built‑in styles to visually differentiate multiple pivot summaries.
// AI Prompts: Add a third pivot table that calculates average revenue per region using the existing data range. | Change the aggregation of the Quantity field from Sum to Count in the first pivot table and refresh it. | Move the second pivot table to a new cell location and update its source range programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsMultiplePivotTables
{
    // This example builds a workbook, fills it with random sales data, defines the entire data block as the source range, and adds two pivot tables to the same sheet—one summarizing total Quantity by Region at F3 and another summarizing total Revenue by Product at F20. Each pivot uses a different built‑in style, is refreshed, calculated, and the workbook is saved as an XLSX file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Header row
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Quantity");
            cells["D1"].PutValue("Revenue");

            // Sample rows
            string[] regions = { "North", "South", "East", "West" };
            string[] products = { "Apple", "Banana", "Cherry" };
            Random rnd = new Random();

            int row = 2;
            for (int i = 0; i < 30; i++)
            {
                cells[row, 0].PutValue(regions[rnd.Next(regions.Length)]);
                cells[row, 1].PutValue(products[rnd.Next(products.Length)]);
                cells[row, 2].PutValue(rnd.Next(1, 100));          // Quantity
                cells[row, 3].PutValue(rnd.Next(100, 1000));      // Revenue
                row++;
            }

            // Define the source data range for pivot tables
            // Using a formula style reference to the whole data block
            string sourceData = $"=Sheet1!{cells.MaxDisplayRange.Address}";

            // -------------------------------------------------
            // First PivotTable: Summarize total Quantity by Region
            // -------------------------------------------------
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex1 = pivots.Add(sourceData, "F3", "QuantityByRegion");
            PivotTable pivot1 = pivots[pivotIndex1];

            // Row field: Region
            pivot1.AddFieldToArea(PivotFieldType.Row, "Region");
            // Data field: Quantity (sum)
            pivot1.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Optional: set a built‑in style
            pivot1.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // -------------------------------------------------
            // Second PivotTable: Summarize total Revenue by Product
            // -------------------------------------------------
            int pivotIndex2 = pivots.Add(sourceData, "F20", "RevenueByProduct");
            PivotTable pivot2 = pivots[pivotIndex2];

            // Row field: Product
            pivot2.AddFieldToArea(PivotFieldType.Row, "Product");
            // Data field: Revenue (sum)
            pivot2.AddFieldToArea(PivotFieldType.Data, "Revenue");

            // Apply a different style
            pivot2.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium12;

            // Refresh and calculate data for both pivot tables
            pivot1.RefreshData();
            pivot1.CalculateData();

            pivot2.RefreshData();
            pivot2.CalculateData();

            // Save the workbook
            workbook.Save("MultiplePivotTables.xlsx");
        }
    }
}
