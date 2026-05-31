using Aspose.Cells;
using Aspose.Cells.Charts;

class AdjustChartPosition
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(40);
        sheet.Cells["B3"].PutValue(60);

        // Add a chart at an initial position (rows 5‑15, columns 0‑5)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Align the chart's top‑left corner with a specific cell range.
        // For example, align the chart to start at cell C2 (row index 1, column index 2)
        // and extend to row 12, column 8 to keep a reasonable size.
        int topRow = 1;      // Row index for the upper‑left corner (C2)
        int leftColumn = 2;  // Column index for the upper‑left corner (C2)
        int bottomRow = 12;  // Row index for the lower‑right corner
        int rightColumn = 8; // Column index for the lower‑right corner

        // Move the chart to the specified cell coordinates
        chart.Move(topRow, leftColumn, bottomRow, rightColumn);

        // Save the workbook with the repositioned chart
        workbook.Save("AdjustedChart.xlsx");
    }
}