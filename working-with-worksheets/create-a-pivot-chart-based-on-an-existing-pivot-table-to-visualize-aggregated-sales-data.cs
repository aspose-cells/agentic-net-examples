using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace PivotChartExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Populate source data (e.g., sales data)
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "SalesData";

            // Header row
            dataSheet.Cells["A1"].PutValue("Region");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            // Sample data
            dataSheet.Cells["A2"].PutValue("North");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("North");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(800);

            dataSheet.Cells["A4"].PutValue("South");
            dataSheet.Cells["B4"].PutValue("Apple");
            dataSheet.Cells["C4"].PutValue(1500);

            dataSheet.Cells["A5"].PutValue("South");
            dataSheet.Cells["B5"].PutValue("Banana");
            dataSheet.Cells["C5"].PutValue(700);

            // -------------------------------------------------
            // 2. Create a worksheet for the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotReport");

            // Define the source range for the pivot table
            string sourceRange = $"=SalesData!A1:C5";

            // Add the pivot table; top‑left corner at A1, name it "SalesPivot"
            int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table:
            // - Row field: Region
            // - Column field: Product
            // - Data field: Sum of Sales
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the pivot data so the chart can read it
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 3. Add a pivot chart based on the created pivot table
            // -------------------------------------------------
            // Add a column chart positioned on the same worksheet
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
            Chart chart = pivotSheet.Charts[chartIndex];

            // Set the chart's data source to the pivot table
            chart.PivotSource = "PivotReport!SalesPivot";

            // Refresh the chart to pull data from the pivot table
            chart.RefreshPivotData();

            // Optional: expose pivot controls on the chart
            PivotOptions options = chart.PivotOptions;
            options.DropZonesVisible = true;
            options.DropZoneCategories = true;
            options.DropZoneSeries = true;
            options.DropZoneData = true;
            options.DropZoneFilter = true;

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            workbook.Save("PivotChart_Sales.xlsx");
        }
    }
}