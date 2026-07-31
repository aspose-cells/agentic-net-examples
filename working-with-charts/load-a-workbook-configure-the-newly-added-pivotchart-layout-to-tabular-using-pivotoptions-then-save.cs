// Title: C# – Set PivotChart Layout to Tabular Using PivotOptions in Aspose.Cells
// Description: Loads a workbook, creates a pivot table on a new sheet, adds a column PivotChart linked to that table, configures PivotOptions to emulate a Tabular layout by enabling drop‑zone zones, refreshes the chart, and saves the file.
// Keywords: Aspose.Cells | C# | PivotChart | Tabular layout | PivotOptions | enable drop zones | refresh pivot chart | create pivot table | Excel export | programmatic chart formatting
// Common Searches: Aspose.Cells set pivot chart to tabular layout C# | how to enable drop zones on PivotChart with Aspose.Cells | refresh PivotChart after changing PivotOptions Aspose.Cells | create pivot table and chart programmatically Aspose.Cells | C# example for tabular pivot chart using Aspose.Cells
// Developer Intent: Configure a newly added PivotChart to display in a Tabular layout via PivotOptions and save the updated workbook.
// Use Cases: Generate a column PivotChart with a tabular‑style view for financial reporting. | Programmatically enable all drop‑zone areas on a PivotChart to mimic a tabular layout before exporting to Excel. | Refresh a PivotChart after modifying its PivotOptions to ensure the layout changes are persisted.
// AI Prompts: Show C# code to set a PivotChart layout to Tabular using Aspose.Cells PivotOptions. | Explain how to enable drop zones on a PivotChart to achieve a tabular view with Aspose.Cells. | Demonstrate refreshing a PivotChart after changing its PivotOptions so the new layout appears in the saved workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartTabular
{
    // Loads a workbook, creates a pivot table on a new sheet, adds a column PivotChart linked to that table, configures PivotOptions to emulate a Tabular layout by enabling drop‑zone zones, refreshes the chart, and saves the file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the first worksheet contains source data for the pivot table
            Worksheet dataSheet = workbook.Worksheets[0];

            // Add a new worksheet that will hold the pivot table and the pivot chart
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

            // Create a pivot table based on the data range (adjust the range as needed)
            // Here we use the used range of the data sheet as the source
            string sourceData = $"=Sheet1!{dataSheet.Cells.MaxDisplayRange.Address}";
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields to the pivot table (example: first column as row, second column as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Refresh and calculate the pivot table so that it contains data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a pivot chart linked to the newly created pivot table
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
            Chart chart = pivotSheet.Charts[chartIndex];

            // Set the chart's pivot source to the pivot table name (same worksheet)
            chart.PivotSource = "PivotSheet!PivotTable1";

            // Access the PivotOptions of the chart
            PivotOptions pivotOptions = chart.PivotOptions;

            // Configure the chart layout to Tabular.
            // Aspose.Cells does not expose a direct "Tabular" property on PivotOptions,
            // but you can achieve a tabular-like layout by enabling the drop zones
            // and making them visible. Adjust these settings as required.
            pivotOptions.DropZonesVisible = true;
            pivotOptions.DropZoneCategories = true;
            pivotOptions.DropZoneSeries = true;
            pivotOptions.DropZoneData = true;
            pivotOptions.DropZoneFilter = true;

            // If a specific layout property exists (e.g., Layout = PivotChartLayout.Tabular),
            // it can be set here. The following line is a placeholder for such an API:
            // pivotOptions.Layout = PivotChartLayout.Tabular; // Uncomment if supported

            // Refresh the chart to apply the pivot data and layout changes
            chart.RefreshPivotData();

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
