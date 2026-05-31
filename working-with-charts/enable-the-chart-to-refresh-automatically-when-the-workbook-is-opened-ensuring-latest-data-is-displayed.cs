using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the source data
            int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

            // Enable automatic refresh of the pivot table when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Add a chart that uses the pivot table as its data source (pivot chart)
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = dataSheet.Charts[chartIndex];
            chart.PivotSource = "Pivot!PivotTable1";

            // Ensure the chart refreshes its data from the pivot table when the workbook opens
            // (the chart will pick up the refreshed pivot data automatically)
            // No additional property is required; setting the pivot's RefreshDataOnOpeningFile is sufficient.

            // Optionally, request Excel to recalculate formulas on open (helps if chart data depends on formulas)
            workbook.Settings.FormulaSettings.CalculateOnOpen = true;

            // Save the workbook
            workbook.Save("ChartAutoRefreshOnOpen.xlsx");
        }
    }
}