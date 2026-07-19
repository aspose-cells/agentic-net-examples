// Title: C# – Aspose.Cells: Column chart with secondary Y‑axis, cell‑based data labels, and centered legend
// Description: Demonstrates how to create a new workbook, populate category and two data series, add a column chart, plot the second series on a secondary Y‑axis, customize its axis range and title, enable cell‑based value and category labels for that series, place the legend at the bottom center, calculate the chart, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells secondary axis chart C# | cell based data labels Aspose.Cells | centered legend Aspose.Cells chart | column chart two Y axes Aspose.Cells | Aspose.Cells chart customization .NET | C# Excel chart secondary Y axis
// Common Searches: Aspose.Cells plot series on secondary Y axis C# | enable data labels from cells for secondary series Aspose.Cells | set legend position bottom center Aspose.Cells chart | create column chart with two Y axes using Aspose.Cells | Aspose.Cells chart axis range and title example
// Developer Intent: Create a column chart with a secondary Y‑axis, show cell‑based labels for the secondary series, and position the legend centrally at the bottom.
// Use Cases: Compare metrics with different units (e.g., revenue vs. units sold) by using a secondary Y‑axis. | Display both the numeric value and the category name directly on the chart for the secondary series. | Achieve a balanced visual layout by centering the legend beneath a chart that occupies most of the worksheet.
// AI Prompts: Generate C# code with Aspose.Cells to build a line chart that uses a secondary Y‑axis, shows cell‑based data labels for the secondary series, and places the legend at the top center. | Explain how to modify the secondary value axis title, minimum, maximum, and major unit after adding a chart in Aspose.Cells. | Provide steps to export a workbook containing a chart with a secondary axis and bottom legend to PDF while preserving all formatting.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // Demonstrates how to create a new workbook, populate category and two data series, add a column chart, plot the second series on a secondary Y‑axis, customize its axis range and title, enable cell‑based value and category labels for that series, place the legend at the bottom center, calculate the chart, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate sample data ----------
                // Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Primary series values (plotted on primary Y axis)
                sheet.Cells["B1"].PutValue("Primary");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                // Secondary series values (plotted on secondary Y axis)
                sheet.Cells["C1"].PutValue("Secondary");
                sheet.Cells["C2"].PutValue(5000);
                sheet.Cells["C3"].PutValue(3000);
                sheet.Cells["C4"].PutValue(1000);

                // ---------- Add a column chart ----------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series
                chart.NSeries.Add("B2:B4", true); // primary series
                chart.NSeries.Add("C2:C4", true); // secondary series

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary Y axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Optional: customize the secondary Y axis (title, range, etc.)
                Axis secValueAxis = chart.SecondValueAxis;
                secValueAxis.Title.Text = "Secondary Axis";
                secValueAxis.MinValue = 0;
                secValueAxis.MaxValue = 6000;
                secValueAxis.MajorUnit = 1000;

                // Enable cell‑based data labels for the secondary series
                chart.NSeries[1].DataLabels.ShowValue = true;
                chart.NSeries[1].DataLabels.ShowCategoryName = true;

                // ---------- Position legend centrally ----------
                // Place legend at the bottom of the chart area
                chart.Legend.Position = LegendPositionType.Bottom;

                // Calculate the chart to ensure all elements are rendered correctly
                chart.Calculate();

                // Save the workbook
                string outputPath = "ChartWithSecondaryAxis_CentralLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
