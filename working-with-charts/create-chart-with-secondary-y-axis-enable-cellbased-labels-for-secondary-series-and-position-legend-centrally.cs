// Title: Create a column chart with a secondary Y‑axis, cell‑derived data labels for the secondary series, and a bottom legend using Aspose.Cells for .NET
// AI Prompts: Generate a new workbook, add sample data, and insert a column chart where the second series is plotted on a secondary Y‑axis with data labels that display both the cell value and its category. | Configure the secondary value axis range (minimum, maximum, major unit) and set the chart legend to appear at the bottom in Aspose.Cells for .NET. | Enable cell‑derived data labels for the secondary series, customize the secondary axis title, and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells .NET create column chart with secondary axis and data labels from cells | how to show category name and value in data labels for secondary series Aspose.Cells | set chart legend position to bottom Aspose.Cells C# example | customize secondary Y axis scale Aspose.Cells column chart | plot second series on secondary axis Aspose.Cells workbook
// Tags: Aspose.Cells column chart dual axis | cell‑derived data labels Aspose.Cells | chart legend bottom placement Aspose.Cells | custom secondary axis scaling Aspose.Cells | plot series on secondary axis Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart with primary and secondary series, plots the secondary series on a secondary Y‑axis, shows cell‑derived value and category labels for that series, customizes the secondary axis range and title, positions the legend at the bottom, and saves the file as ChartWithSecondaryAxis.xlsx.
class ChartWithSecondaryAxis
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Primary");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            sheet.Cells["C1"].PutValue("Secondary");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(1000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series: primary and secondary
            chart.NSeries.Add("B2:B4", true); // primary series
            chart.NSeries.Add("C2:C4", true); // secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Enable cell‑based data labels for the secondary series
            chart.NSeries[1].DataLabels.ShowValue = true;
            chart.NSeries[1].DataLabels.ShowCategoryName = true;

            // Customize the secondary Y axis (optional)
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Secondary Axis";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // Position the legend at the bottom (inside the chart area is default behavior)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Recalculate the chart to apply all settings
            chart.Calculate();

            // Save the workbook
            string outputPath = "ChartWithSecondaryAxis.xlsx";

            // Ensure the directory exists before saving (if a directory is specified)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
