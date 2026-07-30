// Title: C# – Dynamically Set Progress Bar Chart Title in Aspose.Cells Based on Project Phase
// Description: Creates a workbook, adds task data, inserts a column chart as a progress bar, and sets the chart title at runtime using the current project phase. The example also shows how to format the title’s font size and boldness before saving the file.
// Keywords: Aspose.Cells | C# | dynamic chart title | progress bar chart | column chart | Excel chart title runtime | project phase | chart title formatting | GitHub example | Aspose.Cells API
// Common Searches: Aspose.Cells set chart title programmatically C# | dynamic Excel chart title based on variable | change column chart title at runtime Aspose.Cells | format chart title font Aspose.Cells C# | update progress bar chart title with project phase
// Developer Intent: The developer needs to assign a chart title that reflects the current project phase when generating a progress‑bar chart with Aspose.Cells in C#.
// Use Cases: Automated project status reports where each chart title displays the active phase (Design, Implementation, Testing, etc.). | Weekly dashboard workbooks that programmatically update chart titles for multiple projects. | Reusable Excel templates that adjust chart titles based on data pulled from a database or API.
// AI Prompts: Generate C# code using Aspose.Cells to read a project phase from a database and update an existing chart title accordingly. | Show how to apply font size, bold, and color styling to a chart title that is set dynamically in Aspose.Cells. | Provide a loop that iterates through all charts in a workbook and prefixes each title with its corresponding project phase.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartTitle
{
    // Creates a workbook, adds task data, inserts a column chart as a progress bar, and sets the chart title at runtime using the current project phase. The example also shows how to format the title’s font size and boldness before saving the file.
    class Program
    {
        static void Main()
        {
            // Define the current project phase (could be retrieved from elsewhere)
            string currentPhase = "Implementation";

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the progress bar chart
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["A2"].PutValue("Design");
            sheet.Cells["A3"].PutValue("Development");
            sheet.Cells["A4"].PutValue("Testing");

            sheet.Cells["B1"].PutValue("Completion");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(60);
            sheet.Cells["B4"].PutValue(20);

            // Add a column chart (used here as a simple progress representation)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Dynamically set the chart title based on the current project phase
            chart.Title.IsVisible = true;
            chart.Title.Text = $"Project Progress - {currentPhase} Phase";

            // Optionally adjust title appearance
            chart.Title.Font.Size = 14;
            chart.Title.Font.IsBold = true;

            // Save the workbook
            workbook.Save("DynamicChartTitle.xlsx");
        }
    }
}
