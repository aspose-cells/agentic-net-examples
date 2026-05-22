using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class AddHyperlinkToChartLabel
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);

            // Add a label inside the chart area (coordinates are in 1/4000 of chart area)
            Label label = chart.Shapes.AddLabelInChart(1000, 1000, 2000, 4000);
            label.Text = "Visit Aspose";
            label.Font.Color = Color.Blue;
            label.Font.Size = 12;

            // Configure hyperlink for the label (Hyperlink property is read‑only, set its members instead)
            label.Hyperlink.Address = "https://www.aspose.com";
            label.Hyperlink.TextToDisplay = "Aspose Website";
            label.Hyperlink.ScreenTip = "Click to open Aspose site";

            // Save the workbook with the chart label hyperlink
            string outputPath = "ChartLabelWithHyperlink.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}