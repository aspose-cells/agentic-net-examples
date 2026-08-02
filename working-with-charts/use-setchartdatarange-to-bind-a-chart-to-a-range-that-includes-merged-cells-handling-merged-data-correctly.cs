// Title: Bind an Aspose.Cells Chart to Merged Cells Using SetChartDataRange (C#)
// Description: Demonstrates how to create a workbook, merge cells A2:A4 as a single category, populate series data, add a column chart, and bind the chart to the range A1:C4 with SetChartDataRange(true) while disabling PlotVisibleCellsOnly so the merged label is treated as one category. Saves the result as ChartWithMergedCells.xlsx.
// Keywords: Aspose.Cells SetChartDataRange | merged cells chart C# | column chart merged category | PlotVisibleCellsOnly false | Aspose.Cells chart data range | C# Excel chart merged cells | Aspose.Cells example | Excel merged cells chart binding | Aspose.Cells SetChartDataRange true | chart series by columns
// Common Searches: Aspose.Cells bind chart to merged cells | SetChartDataRange with merged cells C# | how to plot merged categories in Aspose.Cells chart | disable PlotVisibleCellsOnly for merged cells | C# example chart merged cells Aspose
// Developer Intent: Bind a chart to a range that contains merged cells and have the chart interpret the merged area as a single category.
// Use Cases: Building financial dashboards where category labels span multiple rows. | Generating sales reports with grouped product categories displayed in charts. | Creating presentation‑ready workbooks where merged headings must appear as one axis label.
// AI Prompts: Show how to switch the chart to plot series by rows while still handling merged cells correctly. | Add axis titles, a legend, and custom colors to the chart after binding it to a merged‑cell range. | Explain the role of PlotVisibleCellsOnly in merged‑cell handling and when it should be set to false.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, merge cells A2:A4 as a single category, populate series data, add a column chart, and bind the chart to the range A1:C4 with SetChartDataRange(true) while disabling PlotVisibleCellsOnly so the merged label is treated as one category. Saves the result as ChartWithMergedCells.xlsx.
class SetChartDataRangeWithMergedCells
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header row.
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Series1");
        cells["C1"].PutValue("Series2");

        // Add category data and merge cells A2:A4 to represent a single merged category.
        cells["A2"].PutValue("Group 1");
        cells["A3"].PutValue(""); // part of the merged area
        cells["A4"].PutValue(""); // part of the merged area
        // Merge A2:A4 (rows 1‑3, column 0). Parameters: firstRow, firstColumn, totalRows, totalColumns.
        cells.Merge(1, 0, 3, 1);

        // Add numeric data for the series.
        cells["B2"].PutValue(10);
        cells["C2"].PutValue(15);
        cells["B3"].PutValue(20);
        cells["C3"].PutValue(25);
        cells["B4"].PutValue(30);
        cells["C4"].PutValue(35);

        // Add a column chart to the worksheet.
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 7);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the chart to the range that includes the merged cells.
        // The second argument (true) indicates that series are plotted by columns.
        chart.SetChartDataRange("A1:C4", true);

        // Ensure the chart considers merged cells as a single category.
        chart.PlotVisibleCellsOnly = false;

        // Save the workbook.
        workbook.Save("ChartWithMergedCells.xlsx");
    }
}
