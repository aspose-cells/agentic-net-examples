using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartTitle
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for a progress bar chart
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["A2"].PutValue("Design");
            sheet.Cells["A3"].PutValue("Development");
            sheet.Cells["A4"].PutValue("Testing");

            sheet.Cells["B1"].PutValue("Progress");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(60);
            sheet.Cells["B4"].PutValue(90);

            // Add a column chart (used as a simple progress bar)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Determine the current project phase (this could come from any source)
            string currentPhase = GetCurrentProjectPhase(); // e.g., "Design", "Development", "Testing"

            // Dynamically set the chart title based on the current phase
            chart.Title.IsVisible = true;
            chart.Title.Text = $"Progress Bar - {currentPhase} Phase";

            // Optional: format the title
            chart.Title.Font.Name = "Arial";
            chart.Title.Font.Size = 14;
            chart.Title.Font.IsBold = true;

            // Save the workbook
            workbook.Save("DynamicChartTitle.xlsx");
        }

        // Mock method to obtain the current project phase.
        // Replace this with real logic as needed.
        static string GetCurrentProjectPhase()
        {
            // For demonstration, we return a static value.
            // In a real scenario, this could read from a database, file, or UI.
            return "Development";
        }
    }
}