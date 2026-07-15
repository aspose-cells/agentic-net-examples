using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Pivot;

namespace AsposeCellsChartAutoRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare source data for the pivot table / chart
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");

            // Sample rows
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B4"].PutValue(30);

            // -------------------------------------------------
            // Add a PivotTable based on the source data
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Value as data

            // Enable automatic refresh of the pivot data when the file is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // -------------------------------------------------
            // Add a chart linked to the pivot table (pivot chart)
            // -------------------------------------------------
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = pivotSheet.Charts[chartIndex];

            // Set the pivot source for the chart
            chart.PivotSource = "Pivot!PivotTable1";

            // Refresh the chart's data from the pivot table (ensures initial data is correct)
            chart.RefreshPivotData();

            // -------------------------------------------------
            // Ensure the chart cache is refreshed when the workbook is saved
            // -------------------------------------------------
            // This option tells Excel to recalculate the chart cache on opening.
            PdfSaveOptions saveOptions = new PdfSaveOptions(); // Using a generic SaveOptions works as well
            saveOptions.RefreshChartCache = true;

            // Additionally, set the workbook's formula settings to calculate on open
            // (useful if the chart depends on formulas)
            workbook.Settings.FormulaSettings.CalculateOnOpen = true;

            // -------------------------------------------------
            // Save the workbook (save rule)
            // -------------------------------------------------
            workbook.Save("ChartAutoRefreshDemo.xlsx", saveOptions);
        }
    }
}