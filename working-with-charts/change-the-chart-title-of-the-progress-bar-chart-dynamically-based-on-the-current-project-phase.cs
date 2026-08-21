// Title: Set a Dynamic Title for a Progress Bar Chart in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills task and completion data, adds a column chart styled as a progress bar, and assigns a title that incorporates a runtime variable (currentPhase). The title is shown, bolded, and sized to 14 pt before the file is saved.
// Keywords: Aspose.Cells | .NET | C# | dynamic chart title | progress bar chart | update Excel chart title at runtime | chart title formatting | column chart Aspose.Cells | project phase label | Excel automation
// Common Searches: Aspose.Cells change chart title dynamically C# | set Excel chart title from variable .NET | progress bar chart title based on project phase | format chart title bold 14 point Aspose.Cells | runtime update of Excel chart title using C#
// Developer Intent: Add a runtime‑driven title to a progress bar chart in an Excel workbook using Aspose.Cells.
// Use Cases: Generate weekly status reports where the chart title reflects the current phase (Design, Development, Testing, etc.). | Build a template that reads the phase name from a database or worksheet cell and automatically refreshes the chart title when data changes. | Create presentation‑ready workbooks with a bold, larger‑font title that highlights the active project phase.
// AI Prompts: Write C# code with Aspose.Cells that sets a chart title from a string variable and formats it as bold 14‑point text. | Show how to bind a chart title to a worksheet cell so the title updates automatically when the cell value changes. | Provide an example that iterates over multiple charts in a workbook and assigns each a title from an array of project phase names.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartDemo
{
    // Creates a workbook, fills task and completion data, adds a column chart styled as a progress bar, and assigns a title that incorporates a runtime variable (currentPhase). The title is shown, bolded, and sized to 14 pt before the file is saved.
    class Program
    {
        static void Main()
        {
            // Define the current project phase (could be retrieved from elsewhere)
            string currentPhase = "Design";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the progress bar chart
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["A2"].PutValue("Planning");
            sheet.Cells["A3"].PutValue("Development");
            sheet.Cells["A4"].PutValue("Testing");

            sheet.Cells["B1"].PutValue("Completion");
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(50);
            sheet.Cells["B4"].PutValue(80);

            // Add a column chart (can be styled as a progress bar)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Dynamically set the chart title based on the current project phase
            chart.Title.IsVisible = true;
            chart.Title.Text = $"Project Progress - {currentPhase} Phase";

            // Optional: adjust title appearance
            chart.Title.Font.Size = 14;
            chart.Title.Font.IsBold = true;

            // Save the workbook
            workbook.Save("ProgressBarChart.xlsx");
        }
    }
}
