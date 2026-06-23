using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class InsertLineChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Populate sample sales data -----
        // Header row
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Product A");
        sheet.Cells["C1"].PutValue("Product B");
        sheet.Cells["D1"].PutValue("Product C");

        // Quarter labels
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        // Sales figures for each product
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["C2"].PutValue(150);
        sheet.Cells["D2"].PutValue(130);

        sheet.Cells["B3"].PutValue(135);
        sheet.Cells["C3"].PutValue(160);
        sheet.Cells["D3"].PutValue(145);

        sheet.Cells["B4"].PutValue(150);
        sheet.Cells["C4"].PutValue(170);
        sheet.Cells["D4"].PutValue(155);

        sheet.Cells["B5"].PutValue(165);
        sheet.Cells["C5"].PutValue(180);
        sheet.Cells["D5"].PutValue(170);

        // ----- Add a line chart -----
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add series data.
        // Data range includes the header row (B1:D5) so that series names are taken from the first row.
        // isVertical = true (series are in columns), checkLabels = true (first row contains series names)
        chart.NSeries.Add("B1:D5", true, true);

        // Set category (X‑axis) data to the quarters column.
        chart.NSeries.CategoryData = "A2:A5";

        // Optional: set chart title and enable legend
        chart.Title.Text = "Quarterly Sales Comparison";
        chart.ShowLegend = true;

        // Save the workbook with the chart
        workbook.Save("LineChartMultipleQuarters.xlsx");
    }
}