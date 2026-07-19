// Title: Set a Progress‑Bar Chart Title Dynamically Based on Project Phase with Aspose.Cells for .NET
// Description: Shows how to build a workbook, add task data, insert a bar chart that acts as a progress bar, and programmatically assign a title that reflects the current project phase at runtime before saving the file.
// Keywords: Aspose.Cells | C# chart title | dynamic chart title | progress bar chart | Excel bar chart .NET | set chart title programmatically | project phase | runtime chart update
// Common Searches: Aspose.Cells change chart title at runtime | C# set Excel chart title dynamically | progress bar chart title variable | update chart title with project phase Aspose | dynamic Excel chart title .NET example
// Developer Intent: The developer needs to assign a chart title that automatically incorporates the current project phase when generating a progress‑bar style chart.
// Use Cases: Generate status‑report workbooks where each progress‑bar chart displays the active project phase in its title. | Automate weekly project dashboards that pull the phase from a database and reflect it in chart titles. | Create templates for multiple projects that insert a bar chart and set its title based on a phase field without manual editing.
// AI Prompts: Write C# code using Aspose.Cells to update an existing chart title with a value read from a database column. | Show how to bind a chart title to a worksheet cell so the title changes when the cell value is edited. | Provide an example that loops through all worksheets in a workbook and sets each chart's title to the corresponding project phase stored in a variable.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, add task data, insert a bar chart that acts as a progress bar, and programmatically assign a title that reflects the current project phase at runtime before saving the file.
class Program
{
    static void Main()
    {
        // Determine the current project phase (replace with real logic as needed)
        string currentPhase = GetCurrentProjectPhase();

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for a progress‑bar style chart
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Completion");
        sheet.Cells["A2"].PutValue("Task 1");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Task 2");
        sheet.Cells["B3"].PutValue(60);
        sheet.Cells["A4"].PutValue("Task 3");
        sheet.Cells["B4"].PutValue(90);

        // Insert a bar chart (used as a progress bar)
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Bind data to the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Dynamically set the chart title based on the current phase
        chart.Title.IsVisible = true;
        chart.Title.Text = $"Project Progress – {currentPhase} Phase";

        // Save the workbook
        workbook.Save("ProgressBarChart.xlsx");
    }

    // Placeholder for obtaining the current project phase.
    // Replace this with actual project‑phase detection logic.
    static string GetCurrentProjectPhase()
    {
        return "Implementation";
    }
}
