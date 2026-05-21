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

            // Sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(10 + i * 5);
            ws.Cells["B3"].PutValue(20 + i * 5);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 8);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";
        }

        // Save each chart as a PNG file named after its worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            for (int c = 0; c < ws.Charts.Count; c++)
            {
                Chart chart = ws.Charts[c];
                string imageFile = $"{ws.Name}_Chart{c + 1}.png"; // PNG format via extension
                chart.ToImage(imageFile); // Uses ToImage(string) overload
                Console.WriteLine($"Chart saved as {imageFile}");
            }
        }

        // Save the workbook (optional)
        workbook.Save("ChartsWorkbook.xlsx", SaveFormat.Xlsx);
    }
}