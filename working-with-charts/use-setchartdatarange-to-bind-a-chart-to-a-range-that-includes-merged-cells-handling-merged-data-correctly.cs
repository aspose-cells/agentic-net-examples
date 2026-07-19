// Title: C# – Bind an Aspose.Cells Chart to Merged Cells with SetChartDataRange
// Description: This example creates a workbook, merges category cells (A2:A3), adds a column chart, and calls SetChartDataRange with the includeMergedCells flag set to true. PlotVisibleCellsOnly is disabled so the chart includes all visible data, and the file is saved as ChartWithMergedCells.xlsx.
// Keywords: Aspose.Cells | SetChartDataRange | merged cells | C# | column chart | includeMergedCells | PlotVisibleCellsOnly | chart data range | Excel merged cells chart
// Common Searches: Aspose.Cells SetChartDataRange merged cells C# | bind chart to merged range Aspose.Cells | include merged cells in chart Aspose.Cells | PlotVisibleCellsOnly false chart merged cells | how to merge cells for chart labels Aspose.Cells
// Developer Intent: Create a chart that correctly reads and displays data from a range containing merged cells.
// Use Cases: Generate a column chart where category labels span multiple rows using merged cells. | Produce a financial report with merged row headers and a chart that reflects the combined categories. | Export an Excel workbook with merged category cells while preserving accurate chart rendering.
// AI Prompts: Show C# code that uses Aspose.Cells SetChartDataRange to bind a chart to a range with merged cells. | Explain the effect of the includeMergedCells flag and PlotVisibleCellsOnly property on chart rendering with merged cells. | Provide a step‑by‑step example of merging cells for row labels and creating a chart that reads those labels correctly in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, merges category cells (A2:A3), adds a column chart, and calls SetChartDataRange with the includeMergedCells flag set to true. PlotVisibleCellsOnly is disabled so the chart includes all visible data, and the file is saved as ChartWithMergedCells.xlsx.
class SetChartDataRangeMergedDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add headers
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");

        // Add data rows
        cells["A2"].PutValue("Group 1");
        cells["A3"].PutValue("Group 1"); // will be merged with A2
        cells["B2"].PutValue(10);
        cells["B3"].PutValue(15);

        cells["A4"].PutValue("Group 2");
        cells["B4"].PutValue(20);

        // Merge the category cells for "Group 1" (A2:A3)
        // Parameters: firstRow (zero‑based), firstColumn, totalRows (1‑based), totalColumns (1‑based)
        cells.Merge(1, 0, 2, 1); // Merges A2:A3

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the chart to a range that includes the merged cells
        // The range covers the header row and all data rows
        chart.SetChartDataRange("A1:B4", true);

        // Ensure the chart plots all visible cells (including merged ones)
        chart.PlotVisibleCellsOnly = false;

        // Save the workbook
        workbook.Save("ChartWithMergedCells.xlsx");
    }
}
