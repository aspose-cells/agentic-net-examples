// Title: Create a radar chart with a transposed data series, enable axis and data labels, and auto‑fit label shapes using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to generate a workbook, adds a radar chart based on a transposed range, turns on radar axis labels, and configures the series to display data labels with values and category names in a rectangular shape that automatically resizes to fit the text. | Show how to set the chart data range by rows, add a series from a column range, enable HasRadarAxisLabels, and customize DataLabels (ShowValue, ShowCategoryName, ShapeType = Rect, IsResizeShapeToFitText = true) in Aspose.Cells.
// Common Searches: Aspose.Cells C# radar chart using transposed range for series | How to show category names and values in radar chart data labels with Aspose.Cells | Auto resize data label shape in Aspose.Cells radar chart | Enable radar axis labels in a .NET workbook using Aspose.Cells | Set chart data range by rows instead of columns Aspose.Cells example
// Tags: radar chart data labels Aspose.Cells | transposed series range Aspose.Cells | auto‑fit label shape Aspose.Cells | enable radar axis labels Aspose.Cells | set chart data range by rows Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsRadarChartDemo
{
    // The example creates a new workbook, fills it with sample category and series data, adds a radar chart using a transposed cell range, activates radar axis labels, and configures data labels to show both values and category names in rectangular shapes that automatically resize to fit the text. The workbook is saved as RadarChartWithDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Category labels
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Cat1");
                sheet.Cells["A3"].PutValue("Cat2");
                sheet.Cells["A4"].PutValue("Cat3");

                // Series values
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(4);
                sheet.Cells["B3"].PutValue(2);
                sheet.Cells["B4"].PutValue(5);

                // Add a radar chart
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (plots by rows)
                chart.SetChartDataRange("A1:B4", false);

                // Add series using the transposed range.
                // The 'true' flag indicates that the first column/row contains series names.
                int seriesIdx = chart.NSeries.Add("B2:B4", true);

                // Enable radar axis labels (category axis labels)
                chart.NSeries[seriesIdx].HasRadarAxisLabels = true;

                // Enable data labels for the series
                Series series = chart.NSeries[seriesIdx];
                series.DataLabels.ShowValue = true;               // Show the numeric values
                series.DataLabels.ShowCategoryName = true;        // Show category names
                series.DataLabels.ShapeType = DataLabelShapeType.Rect; // Rectangle shape for labels
                series.DataLabels.IsResizeShapeToFitText = true;   // Auto‑fit shape to text

                // Recalculate chart layout after modifications
                chart.Calculate();

                // Define output file path
                string outputPath = "RadarChartWithDataLabels.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
