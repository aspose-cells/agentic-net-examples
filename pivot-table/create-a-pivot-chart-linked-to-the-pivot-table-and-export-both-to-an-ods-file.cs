using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Ods;   // OdsSaveOptions namespace

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(2000);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Add a chart and link it to the pivot table (pivot chart)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        // Set the pivot source; the chart becomes a pivot chart
        chart.PivotSource = $"{sheet.Name}!SalesPivot";

        // Refresh chart data from the pivot table
        chart.RefreshPivotData();

        // Save the workbook as ODS, including the pivot table and chart
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.IgnorePivotTables = false; // ensure pivot tables are saved
        workbook.Save("PivotChartDemo.ods", saveOptions);
    }
}