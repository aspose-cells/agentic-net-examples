// Title: Resize a column chart to 400 × 300 points with Aspose.Cells for C#
// Common Searches: Aspose.Cells C# set column chart width to 400 points | How to change chart height to 300 points using Aspose.Cells API | Programmatically adjust Excel chart dimensions in C# with Aspose.Cells | Set ChartObject.Width and Height properties in Aspose.Cells example | Resize chart object in generated XLSX file using Aspose.Cells C#
// Tags: Aspose.Cells chartobject width property | Aspose.Cells chartobject height property | C# resize Excel chart dimensions | column chart size adjustment Aspose.Cells | set chart size points Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartResizeDemo
{
    // The sample creates a workbook, adds data, inserts a column chart, then uses ChartObject.Width = 400 and ChartObject.Height = 300 to resize the chart to 400 × 300 points before saving the file as ResizedChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Resize the chart: width = 400 points, height = 300 points
            // ChartObject represents the shape that contains the chart.
            chart.ChartObject.Width = 400;   // Width in points
            chart.ChartObject.Height = 300;  // Height in points

            // Optional: give the chart a title so the result is obvious
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook to a file
            string outputPath = "ResizedChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Chart resized to 400x300 points and saved to '{outputPath}'.");
        }
    }
}
