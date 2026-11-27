// Title: How to apply a rounded‑rectangle shape to data labels in a line chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set data label shape to rounded rectangle for line chart in C# | C# line chart data labels round rectangle shape Aspose.Cells example | Change data label shape type to RoundRect in Aspose.Cells chart | Set data label position above points and shape to rounded rectangle with Aspose.Cells | DataLabelShapeType.RoundRect usage in Aspose.Cells line chart
// Tags: Aspose.Cells line chart data label shape | DataLabelShapeType.RoundRect C# example | set data label shape Aspose.Cells | line chart data label position above Aspose.Cells | Excel chart rounded rectangle data labels Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLineChartDataLabelShape
{
    // The sample creates a workbook, adds sample data, inserts a line chart, enables data labels for the first series, sets the data label shape to a rounded rectangle via DataLabelShapeType.RoundRect, positions the labels above the points, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the line chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");
                worksheet.Cells["A5"].PutValue("Apr");

                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(15);
                worksheet.Cells["B5"].PutValue(25);

                // Add a line chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 7, 0, 25, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories

                // Access the first series and enable data labels
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Set the shape type of the data labels to rounded rectangle
                series.DataLabels.ShapeType = DataLabelShapeType.RoundRect;

                // Optional: adjust label position if desired
                series.DataLabels.Position = LabelPositionType.Above;

                // Save the workbook to a file
                string outputPath = "LineChartWithRoundedRectDataLabels.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"Data label shape type set to: {series.DataLabels.ShapeType}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
