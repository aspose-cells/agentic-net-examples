using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

class PivotChartDashboard
{
    static void Main()
    {
        // 1. Create a new workbook
        Workbook workbook = new Workbook();

        // 2. Add sample data to the first worksheet (acts as a table)
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Header
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");

        // Sample rows
        dataSheet.Cells["A2"].PutValue("Food");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["A3"].PutValue("Travel");
        dataSheet.Cells["B3"].PutValue(80);
        dataSheet.Cells["A4"].PutValue("Food");
        dataSheet.Cells["B4"].PutValue(150);
        dataSheet.Cells["A5"].PutValue("Utilities");
        dataSheet.Cells["B5"].PutValue(200);

        // 3. Add a worksheet that will hold the PivotTable
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // 4. Define the source data range for the pivot table
        // Use the MaxDisplayRange to get the used range dynamically
        string sourceRange = $"=Data!{dataSheet.Cells.MaxDisplayRange.Address}";

        // 5. Add a PivotTable (using the Add(string, string, string) rule)
        int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // 6. Configure the PivotTable fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");    // Data field (sum)

        // Optional: display in tabular form
        pivotTable.ShowInTabularForm();

        // Refresh and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // 7. Add a dashboard worksheet where the pivot chart will be placed
        Worksheet dashboardSheet = workbook.Worksheets.Add("Dashboard");

        // 8. Add a chart to the dashboard sheet (using the Charts.Add rule)
        // Parameters: ChartType, first row, first column, last row, last column
        int chartIndex = dashboardSheet.Charts.Add(ChartType.Column, 2, 0, 20, 7);
        Chart chart = dashboardSheet.Charts[chartIndex];

        // 9. Link the chart to the pivot table (using Chart.PivotSource property)
        // Since the pivot table is on the "Pivot" sheet, reference it accordingly
        chart.PivotSource = "Pivot!SalesPivot";

        // 10. Refresh the chart to pull data from the pivot table (using RefreshPivotData)
        chart.RefreshPivotData();

        // 11. Save the workbook
        workbook.Save("PivotChartDashboard.xlsx", SaveFormat.Xlsx);
    }
}