// Title: Configure Z‑Axis Tick Mark Spacing to Show Every Second Depth Level in a 3‑D Column Chart (Aspose.Cells for .NET)
// Description: C# sample that creates a workbook, inserts quarterly data, adds a 3‑D column chart, and sets SeriesAxis.TickMarkSpacing to 2 so the Z‑axis (depth/series) displays a tick and label for every other level before saving the file as an .xlsx document.
// Keywords: Aspose.Cells | .NET | C# | 3D column chart | Z axis interval | SeriesAxis.TickMarkSpacing | depth level spacing | chart axis customization | Excel 3D chart Aspose | alternate series labels | tick mark spacing
// Common Searches: Aspose.Cells set Z axis interval 3D chart | SeriesAxis.TickMarkSpacing example C# | display every other depth level in 3D column chart | how to adjust Z axis tick spacing Aspose.Cells | configure depth axis labels Aspose.Cells .NET
// Developer Intent: Apply SeriesAxis.TickMarkSpacing to the Z (series) axis of a 3‑D chart so that only every second depth level is labeled.
// Use Cases: Quarterly sales dashboard where alternating series labels reduce clutter on a 3‑D column chart. | Financial report with dozens of product series, showing only selected depth levels for readability. | Presentation‑ready Excel file that highlights specific depth levels by increasing the Z‑axis interval.
// AI Prompts: Generate C# code using Aspose.Cells to set the Z‑axis tick interval to 3 for a 3‑D bar chart. | Explain the effect of SeriesAxis.TickMarkSpacing on chart rotation and perspective in Aspose.Cells. | Show how to hide Z‑axis labels while keeping tick marks visible in a 3‑D chart with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# sample that creates a workbook, inserts quarterly data, adds a 3‑D column chart, and sets SeriesAxis.TickMarkSpacing to 2 so the Z‑axis (depth/series) displays a tick and label for every other level before saving the file as an .xlsx document.
    public class ConfigureZAxisIntervalDemo
    {
        // Entry point for the application
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

            // Populate sample data for a 3‑D column chart
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
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the Z (depth/series) axis to display every second depth level
            // TickMarkSpacing defines the number of series between tick marks.
            // Setting it to 2 shows a tick (and label) for every second depth level.
            chart.SeriesAxis.TickMarkSpacing = 2;

            // Optional: adjust other 3‑D view properties for better visibility
            chart.Elevation = 30;
            chart.RotationAngle = 20;
            chart.Perspective = 30;

            // Save the workbook
            workbook.Save("ConfigureZAxisIntervalDemo.xlsx");
        }
    }
}
