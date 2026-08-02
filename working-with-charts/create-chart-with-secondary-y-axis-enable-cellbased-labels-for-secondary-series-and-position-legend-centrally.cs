// Title: Aspose.Cells .NET Example: Column Chart with Secondary Y‑Axis, Cell‑Based Data Labels, and Centered Legend (C#)
// Description: A complete C# sample that creates a new workbook, fills cells with categories and two data series, adds a column chart, plots the second series on a secondary Y‑axis with a custom title and range, shows value labels taken from cells, overlays the legend at the bottom and centers it horizontally, then saves the file. Ideal for developers needing dual‑axis charts and precise legend placement using Aspose.Cells for .NET.
// Keywords: Aspose.Cells secondary axis chart C# | column chart dual Y axis Aspose.Cells | cell based data labels Aspose.Cells | centered legend Aspose.Cells .NET | Aspose.Cells chart customization example | GitHub Aspose.Cells chart sample | C# Aspose.Cells chart tutorial
// Common Searches: Aspose.Cells plot series on secondary Y axis | how to add data labels from cells in Aspose.Cells | center legend bottom Aspose.Cells chart | dual axis column chart Aspose.Cells .NET | Aspose.Cells chart example GitHub
// Developer Intent: Generate a column chart where one series uses a secondary Y‑axis, display its values as cell‑based data labels, and place the legend centered at the bottom of the chart.
// Use Cases: Compare monthly sales (primary axis) with profit margin percentages (secondary axis) while showing margin values directly on the chart. | Create a financial dashboard that plots revenue and expense ratios on separate axes, with the ratios labeled from worksheet cells for quick reference. | Design a presentation slide where a dual‑axis chart needs a balanced look, achieved by centering the legend at the bottom of the chart area.
// AI Prompts: Write C# code using Aspose.Cells to add a line series on the secondary Y‑axis and format its data labels as percentages. | Show how to calculate the optimal major unit for a secondary axis based on the data range and apply it in Aspose.Cells. | Provide an Aspose.Cells example that positions the chart legend at the top‑right corner without overlay.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // A complete C# sample that creates a new workbook, fills cells with categories and two data series, adds a column chart, plots the second series on a secondary Y‑axis with a custom title and range, shows value labels taken from cells, overlays the legend at the bottom and centers it horizontally, then saves the file. Ideal for developers needing dual‑axis charts and precise legend placement using Aspose.Cells for .NET.
    public class ChartWithSecondaryAxisAndCentralLegend
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data
                // -------------------------------------------------
                // Primary categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Primary series (plotted on primary Y axis)
                sheet.Cells["B1"].PutValue("Primary");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                // Secondary series (plotted on secondary Y axis)
                sheet.Cells["C1"].PutValue("Secondary");
                sheet.Cells["C2"].PutValue(5000);
                sheet.Cells["C3"].PutValue(3000);
                sheet.Cells["C4"].PutValue(1000);

                // -------------------------------------------------
                // Add a column chart
                // -------------------------------------------------
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIdx];

                // Add the two series
                chart.NSeries.Add("B2:B4", true); // primary series
                chart.NSeries.Add("C2:C4", true); // secondary series
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Configure secondary Y axis for the second series
                // -------------------------------------------------
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Optional: give the secondary axis a title and range
                Axis secAxis = chart.SecondValueAxis;
                secAxis.Title.Text = "Secondary Axis";
                secAxis.MinValue = 0;
                secAxis.MaxValue = 6000;
                secAxis.MajorUnit = 1000;

                // -------------------------------------------------
                // Enable cell‑based data labels for the secondary series
                // -------------------------------------------------
                chart.NSeries[1].DataLabels.ShowValue = true;

                // -------------------------------------------------
                // Position the legend at the centre of the chart area
                // -------------------------------------------------
                chart.Legend.IsOverLay = true;
                chart.Legend.Position = LegendPositionType.Bottom;

                // Center the legend horizontally by using the chart area width
                chart.Legend.X = (chart.ChartArea.Width - chart.Legend.Width) / 2;

                // -------------------------------------------------
                // Recalculate the chart to apply all settings
                // -------------------------------------------------
                chart.Calculate();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "ChartWithSecondaryAxisAndCentralLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
            ChartWithSecondaryAxisAndCentralLegend.Run();
        }
    }
}
