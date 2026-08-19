// Title: Create a Transparent Background Series for a Stacked Column Progress Bar Chart with Aspose.Cells for .NET (C#)
// Description: This example builds an Excel workbook, adds task names and progress values, inserts a stacked column chart, defines a visible progress series and a background series, sets the background series Area.Transparency to 1.0 (100 % transparent) to simulate a progress‑bar effect, optionally colors the progress series, and saves the file as ProgressBarChart.xlsx.
// Keywords: Aspose.Cells C# transparent series | stacked column progress bar Aspose.Cells | chart series transparency .NET | Excel progress bar chart code | Aspose.Cells chart fill format | C# Excel chart background invisible
// Common Searches: Aspose.Cells make chart series invisible C# | transparent background series stacked column chart Aspose.Cells | progress bar chart using Aspose.Cells .NET | set series area transparency Aspose.Cells | create Excel progress bar with Aspose.Cells
// Developer Intent: Hide the background series of a stacked column chart by applying full transparency so only the progress portion remains visible.
// Use Cases: Display task completion percentages as compact progress bars in financial or project reports. | Design clean dashboards where only the filled portion of each bar is shown. | Generate printable Excel sheets with minimalist progress indicators for status updates.
// AI Prompts: Generate C# code with Aspose.Cells that creates a stacked column chart and sets the second series Area.Transparency to 1.0 for a progress‑bar effect. | Explain how to configure series fill and transparency in Aspose.Cells to simulate a progress bar chart. | Provide step‑by‑step instructions to add a visible progress series and an invisible background series to a chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace ProgressBarChartDemo
{
    // This example builds an Excel workbook, adds task names and progress values, inserts a stacked column chart, defines a visible progress series and a background series, sets the background series Area.Transparency to 1.0 (100 % transparent) to simulate a progress‑bar effect, optionally colors the progress series, and saves the file as ProgressBarChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for the progress bar chart
                // Column A – categories, Column B – actual progress values,
                // Column C – invisible series (used to create the background of the bar)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["A4"].PutValue("Task 3");

                sheet.Cells["B1"].PutValue("Progress");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(60);
                sheet.Cells["B4"].PutValue(90);

                sheet.Cells["C1"].PutValue("Background");
                // Background values are the maximum value (e.g., 100) for each task
                sheet.Cells["C2"].PutValue(100);
                sheet.Cells["C3"].PutValue(100);
                sheet.Cells["C4"].PutValue(100);

                // Add a stacked column chart (used as a progress bar)
                int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add the visible progress series (first series)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Progress";

                // Add the invisible background series (second series)
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Background";

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // Make the background series transparent so only the progress part is visible
                Series backgroundSeries = chart.NSeries[1];
                if (backgroundSeries.Area != null)
                {
                    backgroundSeries.Area.Transparency = 1.0; // 100% transparent
                }

                // Optional: give the progress series a solid fill color
                Series progressSeries = chart.NSeries[0];
                if (progressSeries.Area != null && progressSeries.Area.FillFormat != null && progressSeries.Area.FillFormat.SolidFill != null)
                {
                    progressSeries.Area.FillFormat.SolidFill.Color = Color.Green;
                }

                // Save the workbook
                workbook.Save("ProgressBarChart.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
