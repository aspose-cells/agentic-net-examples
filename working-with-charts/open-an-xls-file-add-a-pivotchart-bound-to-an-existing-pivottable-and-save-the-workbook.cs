// Title: Create and bind a PivotChart to an existing PivotTable in an XLS workbook using Aspose.Cells for .NET
// Description: The sample loads a legacy XLS file, opens its first sheet, adds a column chart, connects the chart to the PivotTable called PivotTable1, refreshes both chart and pivot data, and writes the result as an XLSX workbook.
// Keywords: Aspose.Cells | C# PivotChart | bind chart to PivotTable | refresh pivot data | add chart to XLS | convert XLS to XLSX | Aspose.Cells Chart API | PivotTable chart binding | Excel automation .NET
// Common Searches: Aspose.Cells add PivotChart to existing PivotTable | C# bind chart to PivotTable Aspose | Refresh PivotChart data after creation Aspose.Cells | Convert legacy XLS with pivot tables to XLSX using Aspose | How to create chart from PivotTable in .NET
// Developer Intent: Insert a visual chart that reflects a pre‑existing PivotTable and export the workbook in the modern format.
// Use Cases: Produce a column chart for sales figures already summarized in a PivotTable without recreating the source data. | Ensure all pivot tables are up‑to‑date after chart insertion before delivering the file to end users. | Migrate old XLS workbooks that contain pivot tables to XLSX while adding a ready‑to‑use chart.
// AI Prompts: Generate C# code with Aspose.Cells that adds a column PivotChart linked to a PivotTable named 'PivotTable1' in an existing XLS file and saves it as XLSX. | Explain the steps to refresh pivot tables and chart data after creating a PivotChart with Aspose.Cells. | Provide a robust C# example that checks for the presence of the specified PivotTable before binding a chart, and handles the error gracefully.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// The sample loads a legacy XLS file, opens its first sheet, adds a column chart, connects the chart to the PivotTable called PivotTable1, refreshes both chart and pivot data, and writes the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Get the worksheet that contains the existing PivotTable
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a new chart (e.g., Column chart) to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Bind the chart to the existing PivotTable (assumed name "PivotTable1")
        chart.PivotSource = $"{worksheet.Name}!PivotTable1";

        // Refresh the chart's data from the pivot table
        chart.RefreshPivotData();

        // Refresh all pivot tables in the worksheet (optional, ensures data consistency)
        worksheet.RefreshPivotTables();

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
