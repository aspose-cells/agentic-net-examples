// Title: Resize chart data label shapes with large markers using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a line chart, binds a series to sample data, sets 30‑point circular markers, enables data labels above each point, disables auto‑fit, and applies a custom 80 × 30‑point size to every label before recalculating and saving the file.
// Keywords: Aspose.Cells | C# chart customization | line chart data labels | custom label width height | .NET Excel chart | large marker size | disable auto fit data labels | resize chart label shape | Excel export Aspose.Cells
// Common Searches: Aspose.Cells set custom size for chart data labels | Resize data label shapes in a line chart C# | Disable auto‑fit for Excel chart labels Aspose | How to change marker size and label dimensions in Aspose.Cells | Custom width and height for chart point labels .NET
// Developer Intent: Apply a fixed width and height to each data label shape after enabling large markers on a line chart.
// Use Cases: Generate Excel reports where data labels must stay uniform regardless of their text length. | Design charts with oversized markers while keeping label dimensions consistent to avoid overlap. | Create a reusable .NET routine that standardizes label size for all points in a series.
// AI Prompts: Show C# code to disable auto‑fit and set a custom 80 × 30‑point size for each data label in an Aspose.Cells line chart with large markers. | Provide an Aspose.Cells example that resizes data label shapes for all points in a series. | Explain how to recalculate an Aspose.Cells chart after modifying data label dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a line chart, binds a series to sample data, sets 30‑point circular markers, enables data labels above each point, disables auto‑fit, and applies a custom 80 × 30‑point size to every label before recalculating and saving the file.
    class ResizeDataLabelShapes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the line chart
                worksheet.Cells["A1"].PutValue("X");
                worksheet.Cells["B1"].PutValue("Y");
                for (int i = 0; i < 5; i++)
                {
                    worksheet.Cells[i + 2, 0].PutValue(i + 1);          // X values
                    worksheet.Cells[i + 2, 1].PutValue((i + 1) * 10); // Y values
                }

                // Add a line chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Add a series and bind it to the data range
                int seriesIndex = chart.NSeries.Add("B2:B6", true);
                Series series = chart.NSeries[seriesIndex];
                series.XValues = "A2:A6";

                // Enable markers and set a large marker size
                series.Marker.MarkerStyle = ChartMarkerType.Circle;
                series.Marker.MarkerSize = 30; // size in points (large)

                // Enable data labels for the series
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Above;

                // Resize each data label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Disable auto‑fit so custom dimensions are applied
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set custom width and height for the data label shape (in points)
                    point.DataLabels.Width = 80;
                    point.DataLabels.Height = 30;
                }

                // Recalculate the chart to apply the changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "ResizeDataLabelShapes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ResizeDataLabelShapes.Run();
        }
    }
}
