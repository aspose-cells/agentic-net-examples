using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SaveChartImages
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data and a chart to each worksheet
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            ws.Name = $"Sheet{i + 1}";

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(10 + i * 5);
            ws.Cells["B3"].PutValue(20 + i * 5);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Save the chart as a PNG file named after the worksheet
            string imageFileName = $"{ws.Name}.png";
            chart.ToImage(imageFileName); // Extension determines PNG format
        }

        // Save the workbook to disk
        workbook.Save("ChartsWorkbook.xlsx");
    }
}