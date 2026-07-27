// Title: C# – Hide Tick Marks on the Secondary Y‑Axis in a Mixed Chart with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a mixed column chart with primary and secondary series, plots the second series on the secondary Y‑axis, and removes both major and minor tick marks by setting the axis's TickMarkType to None. It also shows how to hide the axis line before saving the file.
// Keywords: Aspose.Cells hide secondary Y axis tick marks | C# Aspose.Cells secondary axis formatting | remove major tick mark Aspose.Cells | remove minor tick mark Aspose.Cells | mixed chart secondary axis Aspose.Cells | TickMarkType.None C# | Aspose.Cells chart axis visibility | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells hide secondary axis tick marks C# | How to remove tick marks from secondary Y axis in Aspose.Cells | C# mixed chart secondary axis formatting Aspose.Cells | Set TickMarkType.None for secondary axis Aspose.Cells | Hide axis line secondary value axis Aspose.Cells
// Developer Intent: Remove major and minor tick marks (and optionally the axis line) from the secondary Y‑axis of a mixed chart using Aspose.Cells for .NET.
// Use Cases: Generate clean mixed column‑line charts where the secondary axis shows large values without distracting tick marks. | Create Excel reports that emphasize data series while keeping the secondary axis visually minimal. | Customize chart appearance programmatically by hiding secondary axis tick marks and line while retaining the plotted series.
// AI Prompts: Show C# code that hides major and minor tick marks on the secondary Y‑axis of an Aspose.Cells mixed chart. | Provide an Aspose.Cells for .NET example that plots a series on the secondary axis, removes its tick marks, optionally hides the axis line, and saves the workbook. | Explain how to set TickMarkType.None and control axis visibility for the secondary value axis in Aspose.Cells charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a mixed column chart with primary and secondary series, plots the second series on the secondary Y‑axis, and removes both major and minor tick marks by setting the axis's TickMarkType to None. It also shows how to hide the axis line before saving the file.
    public class HideSecondaryYAxisTickMarks
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Primary Series");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["B4"].PutValue(300);

            worksheet.Cells["C1"].PutValue("Secondary Series");
            worksheet.Cells["C2"].PutValue(5000);
            worksheet.Cells["C3"].PutValue(3000);
            worksheet.Cells["C4"].PutValue(1000);

            // Add a mixed chart (column chart for illustration)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Add the two series
            chart.NSeries.Add("B2:B4", true); // Primary series
            chart.NSeries.Add("C2:C4", true); // Secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Access the secondary Y axis
            Axis secondaryValueAxis = chart.SecondValueAxis;

            // Hide major and minor tick marks on the secondary Y axis
            secondaryValueAxis.MajorTickMark = TickMarkType.None;
            secondaryValueAxis.MinorTickMark = TickMarkType.None;

            // Optionally, hide the axis line
            // secondaryValueAxis.IsVisible = false;

            // Save the workbook
            workbook.Save("HideSecondaryYAxisTickMarks.xlsx");
        }
    }
}
