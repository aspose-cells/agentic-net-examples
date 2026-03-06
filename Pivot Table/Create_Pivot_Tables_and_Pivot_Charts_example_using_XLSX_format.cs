using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a new workbook and add sample data for the pivot
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Region");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue("North");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Apple");
            dataSheet.Cells["B3"].PutValue("South");
            dataSheet.Cells["C3"].PutValue(1500);

            dataSheet.Cells["A4"].PutValue("Banana");
            dataSheet.Cells["B4"].PutValue("North");
            dataSheet.Cells["C4"].PutValue(800);

            dataSheet.Cells["A5"].PutValue("Banana");
            dataSheet.Cells["B5"].PutValue("South");
            dataSheet.Cells["C5"].PutValue(950);

            dataSheet.Cells["A6"].PutValue("Cherry");
            dataSheet.Cells["B6"].PutValue("North");
            dataSheet.Cells["C6"].PutValue(400);

            dataSheet.Cells["A7"].PutValue("Cherry");
            dataSheet.Cells["B7"].PutValue("South");
            dataSheet.Cells["C7"].PutValue(600);

            // 2. Add a worksheet that will contain the PivotTable
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");
            string sourceData = "Data!A1:C7";
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // 3. Add a worksheet that will contain the PivotChart
            Worksheet chartSheet = workbook.Worksheets.Add("PivotChart");
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 0, 0, 14, 7);
            Chart chart = chartSheet.Charts[chartIndex];

            // Link the chart to the pivot table
            chart.PivotSource = pivotTable.Name;

            // Optional: set chart title and style
            chart.Title.Text = "Sales by Product and Region";
            chart.Title.Font.IsBold = true;
            chart.PlotArea.Border.IsVisible = true;

            // 4. Save the workbook
            workbook.Save("PivotTableAndChartExample.xlsx");
        }
    }
}