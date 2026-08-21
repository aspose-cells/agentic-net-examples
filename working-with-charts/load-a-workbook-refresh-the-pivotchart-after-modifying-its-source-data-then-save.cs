// Title: Refresh PivotChart After Modifying Source Data – Aspose.Cells C# Example
// Description: Load an Excel workbook with Aspose.Cells, change source cells, call RefreshPivotTables, iterate through charts with a PivotSource and invoke RefreshPivotData, then save the file so the PivotChart reflects the new values.
// Keywords: Aspose.Cells C# refresh PivotChart | RefreshPivotData Aspose.Cells | RefreshPivotTables programmatically | update pivot chart source data .NET | PivotSource property chart Aspose | modify Excel cells and refresh pivot chart | Aspose.Cells example refresh chart
// Common Searches: how to refresh a PivotChart after changing source data using Aspose.Cells | Aspose.Cells C# refresh all pivot tables and charts | programmatically update Excel cells and refresh PivotChart | RefreshPivotData example Aspose.Cells | refresh pivot chart in .NET workbook
// Developer Intent: Update source cells, refresh associated pivot tables and charts, and save the workbook with the latest data.
// Use Cases: Automate monthly sales report generation by editing source values and instantly updating pivot charts. | Batch‑process multiple workbooks to correct data errors and ensure every PivotChart displays the corrected figures. | Build a real‑time dashboard where cell edits trigger automatic refresh of pivot tables and charts before publishing.
// AI Prompts: Write C# code with Aspose.Cells that changes specific cell values, refreshes all pivot tables, and refreshes only charts that have a PivotSource. | Explain how Chart.PivotSource is used to identify pivot charts that need RefreshPivotData in Aspose.Cells. | Suggest robust error‑handling patterns for RefreshPivotData and RefreshPivotTables calls in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an Excel workbook with Aspose.Cells, change source cells, call RefreshPivotTables, iterate through charts with a PivotSource and invoke RefreshPivotData, then save the file so the PivotChart reflects the new values.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the pivot table and pivot chart
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the source data for the pivot table is on the first worksheet
        Worksheet dataSheet = workbook.Worksheets[0];

        // Modify the source data that the pivot table/chart depends on
        dataSheet.Cells["B2"].PutValue(1500);
        dataSheet.Cells["B3"].PutValue(2500);

        // Refresh all pivot tables in the workbook after the data change
        workbook.Worksheets.RefreshPivotTables();

        // Refresh all pivot charts (charts that have a PivotSource defined)
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart chart in ws.Charts)
            {
                if (!string.IsNullOrEmpty(chart.PivotSource))
                {
                    chart.RefreshPivotData();
                }
            }
        }

        // Save the workbook with the refreshed pivot tables and charts
        workbook.Save("output.xlsx");
    }
}
