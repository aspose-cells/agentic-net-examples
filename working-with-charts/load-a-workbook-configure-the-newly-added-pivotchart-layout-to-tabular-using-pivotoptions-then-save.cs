// Title: Aspose.Cells for .NET – Build a Tabular‑Layout PivotChart and Apply PivotOptions
// Description: Load a workbook, add sample data, create a PivotTable in tabular form, link a PivotChart, configure PivotOptions (drop zones), refresh the chart, and save the file.
// Keywords: Aspose.Cells PivotChart tabular | PivotOptions C# | create PivotTable programmatically | refresh PivotChart data | save workbook with chart
// Common Searches: Aspose.Cells set PivotChart to tabular layout | C# configure PivotOptions for a PivotChart | how to refresh PivotChart after PivotTable changes | programmatic PivotTable and PivotChart creation .NET
// Developer Intent: Generate an Excel workbook that contains a tabular‑style PivotTable and a linked PivotChart with customized PivotOptions, then save the result.
// Use Cases: Automated sales summary that shows categories in tabular form with a column chart for executive review. | Financial dashboard where users can drag fields in visible drop zones to reshape the chart after opening the file. | Batch creation of reporting workbooks that require a refreshed PivotChart reflecting the latest calculated data.
// AI Prompts: Give a C# example that switches the PivotChart layout from Tabular to Outline while preserving existing PivotOptions. | Show how to add a chart title and enable data labels on the PivotChart after calling RefreshPivotData. | Explain how to export the generated PivotChart as a PNG image using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Load a workbook, add sample data, create a PivotTable in tabular form, link a PivotChart, configure PivotOptions (drop zones), refresh the chart, and save the file.
class PivotChartTabularLayout
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Use the first worksheet as the source data sheet
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Ensure there is some sample data (if the sheet is empty)
        if (dataSheet.Cells.MaxDataRow == 0 && dataSheet.Cells.MaxDataColumn == 0)
        {
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(40);
        }

        // Add a new worksheet to host the PivotTable and PivotChart
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create a PivotTable based on the source data
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B5", "A1", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Layout the PivotTable in Tabular form
        pivotTable.ShowInTabularForm();

        // Refresh and calculate the PivotTable so that the chart can use up‑to‑date data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Add a PivotChart linked to the created PivotTable
        int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
        Chart chart = pivotSheet.Charts[chartIndex];
        chart.PivotSource = "Pivot!PivotTable1";

        // Configure PivotOptions for the chart (example: make drop zones visible)
        PivotOptions pivotOptions = chart.PivotOptions;
        pivotOptions.DropZonesVisible = true;
        pivotOptions.DropZoneCategories = true;
        pivotOptions.DropZoneSeries = true;
        pivotOptions.DropZoneData = true;
        pivotOptions.DropZoneFilter = true;

        // Refresh chart data from the PivotTable
        chart.RefreshPivotData();

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}
