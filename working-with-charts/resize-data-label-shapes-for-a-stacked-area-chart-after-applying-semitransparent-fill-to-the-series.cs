// Title: Resize data label shapes in a stacked area chart with semi‑transparent fill – Aspose.Cells for .NET
// Description: Creates a workbook, adds an AreaStacked chart, applies a 50 % transparent blue fill to the first series, enables data labels, disables auto‑fit, sets each label width to 60 points, recalculates the chart, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells resize data label | stacked area chart label width | semi transparent series fill | disable data label auto fit | C# chart customization | Excel chart label size | Aspose.Cells chart styling
// Common Searches: how to set fixed width for chart data labels Aspose.Cells | apply transparent fill to a series and resize labels in .NET | disable auto‑fit of data label shapes in Excel chart C# | customize stacked area chart label size Aspose.Cells
// Developer Intent: Set a fixed width for each data label shape in a stacked area chart after applying a semi‑transparent fill to the series.
// Use Cases: Generate Excel reports where overlapping area series need visual distinction while keeping label dimensions consistent. | Create dashboards with uniform data label sizes regardless of the displayed value. | Automate chart styling in bulk‑exported workbooks to improve readability and branding.
// AI Prompts: Show C# code that applies a 50 % transparent fill to a chart series and then fixes the width of all data label shapes using Aspose.Cells. | Provide an example that disables auto‑fit for data labels and sets a custom width in a stacked area chart with Aspose.Cells for .NET. | Explain how to recalculate a chart after changing series transparency and data label dimensions in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds an AreaStacked chart, applies a 50 % transparent blue fill to the first series, enables data labels, disables auto‑fit, sets each label width to 60 points, recalculates the chart, and saves the file as an Excel workbook.
    public class ResizeDataLabelShapesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a stacked area chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["B3"].PutValue(50);
                worksheet.Cells["B4"].PutValue(40);

                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(20);
                worksheet.Cells["C3"].PutValue(35);
                worksheet.Cells["C4"].PutValue(45);

                // Add a stacked area chart (use AreaStacked enum value)
                int chartIndex = worksheet.Charts.Add(ChartType.AreaStacked, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Set data range for the chart (both series)
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series
                Series series = chart.NSeries[0];

                // Apply semi‑transparent fill to the series
                series.Area.FillFormat.SolidFill.Color = Color.Blue; // solid blue fill
                series.Area.Transparency = 0.5;                      // 50 % transparent

                // Enable data labels for the series
                series.DataLabels.ShowValue = true;

                // Resize each data label shape (disable auto‑fit and set a custom width)
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.IsResizeShapeToFitText = false; // prevent auto‑resize
                    point.DataLabels.Width = 60;                     // custom width (points)
                }

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "ResizeDataLabelShapesStackedArea.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesDemo.Run();
        }
    }
}
