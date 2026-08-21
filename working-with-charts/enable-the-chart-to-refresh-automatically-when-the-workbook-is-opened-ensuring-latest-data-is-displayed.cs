// Title: Auto‑Refresh Pivot Chart on Workbook Open with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add source data, build a pivot table, enable RefreshDataOnOpeningFile, link a pivot chart, and set CalculateOnOpen so Excel automatically updates the chart each time the file is opened.
// Keywords: Aspose.Cells | C# | .NET | pivot chart auto refresh | RefreshDataOnOpeningFile | CalculateOnOpen | Excel chart update on open | pivot table refresh | Aspose.Cells chart example
// Common Searches: Aspose.Cells auto refresh pivot chart on open | RefreshDataOnOpeningFile property C# | calculate formulas on workbook open Aspose.Cells | pivot chart update automatically Excel | how to enable chart refresh when opening file using Aspose.Cells
// Developer Intent: Configure a workbook so its pivot chart refreshes automatically when the file is opened.
// Use Cases: Generate a report workbook that always shows the latest data without manual refresh. | Create dashboards where pivot charts reflect real‑time changes in source tables. | Ensure formulas influencing chart data are recalculated on open for accurate visuals.
// AI Prompts: Provide C# code with Aspose.Cells that creates a pivot chart and sets it to refresh automatically on workbook open. | Show how to use RefreshDataOnOpeningFile and CalculateOnOpen properties for a workbook containing a pivot chart. | Explain the steps Aspose.Cells takes to enable automatic chart refresh when an Excel file is opened.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoRefreshDemo
{
    // Demonstrates how to create a workbook, add source data, build a pivot table, enable RefreshDataOnOpeningFile, link a pivot chart, and set CalculateOnOpen so Excel automatically updates the chart each time the file is opened.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate source data for the chart/pivot
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the source data
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value

            // Ensure the pivot table refreshes its data when the file is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Create a chart that uses the pivot table as its source (pivot chart)
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = pivotSheet.Charts[chartIndex];
            chart.PivotSource = "Pivot!PivotTable1";

            // Optional: also request Excel to recalculate formulas on open
            workbook.Settings.FormulaSettings.CalculateOnOpen = true;

            // Save the workbook (Excel will refresh the pivot data and chart on open)
            workbook.Save("ChartAutoRefreshOnOpen.xlsx");
        }
    }
}
