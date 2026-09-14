// Title: How to display every other depth level on a 3‑D column chart by setting the Z‑axis interval with Aspose.Cells for .NET (C#)
// AI Prompts: Create a 3‑D column chart in a workbook and set the series (Z) axis TickMarkSpacing to 2 using Aspose.Cells for .NET. | Adjust the depth axis of an Aspose.Cells chart so that only alternate depth labels are shown. | Save the workbook after configuring the Z‑axis interval for a 3‑D chart in a C# application.
// Common Searches: Aspose.Cells C# set Z axis tick spacing for 3D column chart | Show every second depth level in Excel 3D chart using Aspose.Cells | How to configure series axis interval in a 3‑D chart with Aspose.Cells .NET
// Tags: Aspose.Cells TickMarkSpacing on series axis C# | 3D column chart depth axis interval Aspose.Cells | configure Z axis label spacing .NET Excel chart | Aspose.Cells chart axis customization example | C# Excel 3D chart depth level display

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The sample creates a workbook, adds data, inserts a 3‑D column chart, sets the series (Z) axis TickMarkSpacing to 2 so that only every second depth level is shown, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for a 3‑D column chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");

                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(150);
                worksheet.Cells["B4"].PutValue(180);

                worksheet.Cells["C1"].PutValue("Series 2");
                worksheet.Cells["C2"].PutValue(90);
                worksheet.Cells["C3"].PutValue(110);
                worksheet.Cells["C4"].PutValue(130);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the depth (Z) axis to show every second depth level
                // The depth axis corresponds to the series axis in a 3‑D chart.
                // Setting TickMarkSpacing to 2 makes a tick (and thus a label) appear every second series.
                Axis seriesAxis = chart.SeriesAxis;
                seriesAxis.TickMarkSpacing = 2;

                // Prepare output directory
                string outputPath = "ZAxisIntervalEverySecondDepthLevel.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
