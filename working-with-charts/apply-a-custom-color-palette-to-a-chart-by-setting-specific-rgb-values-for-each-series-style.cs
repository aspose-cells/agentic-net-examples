using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CustomChartPalette
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Populate data: categories in column A, two series in columns B and C
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("Jan");
        ws.Cells["A3"].PutValue("Feb");
        ws.Cells["A4"].PutValue("Mar");

        ws.Cells["B1"].PutValue("Series1");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["B3"].PutValue(20);
        ws.Cells["B4"].PutValue(30);

        ws.Cells["C1"].PutValue("Series2");
        ws.Cells["C2"].PutValue(15);
        ws.Cells["C3"].PutValue(25);
        ws.Cells["C4"].PutValue(35);

        // Add a column chart to the worksheet
        int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = ws.Charts[chartIdx];

        // Set the data range for the series and categories
        chart.NSeries.Add("B1:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Define custom RGB colors for each series
        Color[] customColors = new Color[]
        {
            Color.FromArgb(79, 129, 189),   // Color for Series1
            Color.FromArgb(192, 80, 77)    // Color for Series2
        };

        // Apply the custom colors to each series' area
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            chart.NSeries[i].Area.ForegroundColor = customColors[i];
            chart.NSeries[i].Area.Formatting = FormattingType.Custom;
        }

        // Save the workbook with the customized chart
        workbook.Save("CustomPaletteChart.xlsx", SaveFormat.Xlsx);
    }
}