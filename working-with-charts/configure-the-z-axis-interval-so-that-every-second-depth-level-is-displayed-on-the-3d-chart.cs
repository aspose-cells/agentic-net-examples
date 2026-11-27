// Title: Set Z‑axis interval to display every second depth level in a 3‑D column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Z axis major unit for 3D column chart | How to show every other depth level on a 3D chart with Aspose.Cells .NET | Configure series axis tick label spacing in Aspose.Cells C# example | Disable automatic major unit for Z axis in Aspose.Cells chart | Aspose.Cells 3D chart series axis interval configuration tutorial
// Tags: Aspose.Cells series axis major unit configuration | Aspose.Cells 3D column chart Z axis interval | C# set tick label spacing Aspose.Cells | Aspose.Cells disable automatic major unit chart axis | Aspose.Cells depth axis interval setting | Aspose.Cells chart series axis customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ZAxisInterval3DChart
{
    // The example creates a workbook, populates depth categories and two data series, adds a 3‑D column chart, assigns the depth column as the category axis, and then accesses the chart's SeriesAxis (the Z‑axis). Automatic major unit calculation is turned off, the major unit and tick label spacing are set to 2, and the chart is recalculated so that only every second depth level is labeled. The workbook is saved as ZAxisInterval3DChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate sample data.
                // Row 1 : Header (Depth, Series 1, Series 2)
                // Column A : Depth (Z‑axis) categories
                // Columns B & C : Values for two series
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Depth");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");

                // Depth levels (0‑5)
                for (int i = 0; i <= 5; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(i);               // Depth (Z‑axis)
                    sheet.Cells[i + 1, 1].PutValue(10 + i * 5);     // Series 1 values
                    sheet.Cells[i + 1, 2].PutValue(20 + i * 3);     // Series 2 values
                }

                // ------------------------------------------------------------
                // Add a 3‑D column chart.
                // Placed starting at row 8, column 0 and occupies rows 8‑30, columns 0‑10.
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 8, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                // Category (X‑axis) = Depth column (A2:A7)
                // Values = Series columns (B2:C7)
                chart.NSeries.Add("B2:C7", true);
                chart.NSeries.CategoryData = "A2:A7";

                // ------------------------------------------------------------
                // Configure the Z‑axis (depth/series axis) so that every second depth level is displayed.
                // The series axis represents the Z‑axis in a 3‑D chart.
                // ------------------------------------------------------------
                Axis depthAxis = chart.SeriesAxis; // Use SeriesAxis for the Z‑axis

                // Disable automatic major unit calculation.
                depthAxis.IsAutomaticMajorUnit = false;

                // Set the interval between major tick marks to 2.
                depthAxis.MajorUnit = 2;

                // Show labels only for every second depth level.
                depthAxis.TickLabelSpacing = 2;
                depthAxis.IsAutoTickLabelSpacing = false;

                // Recalculate the chart to apply axis settings.
                chart.Calculate();

                // ------------------------------------------------------------
                // Save the workbook.
                // ------------------------------------------------------------
                string outputPath = "ZAxisInterval3DChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the 3‑D chart:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
