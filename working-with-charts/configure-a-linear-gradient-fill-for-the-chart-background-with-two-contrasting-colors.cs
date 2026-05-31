using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Configure the plot area background with a linear two‑color gradient
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;
        plotAreaFill.FillType = FillType.Gradient; // Enable gradient fill
        // Apply a horizontal (linear) gradient from DarkBlue to Yellow
        plotAreaFill.SetTwoColorGradient(
            Color.DarkBlue,   // First color
            Color.Yellow,    // Second color
            GradientStyleType.Horizontal, // Linear gradient direction
            1);               // Variant (1‑4)

        // Save the workbook with the chart
        workbook.Save("ChartBackgroundGradient.xlsx");
    }
}