// Title: Create Column, Pie, and Line Charts on a Single Worksheet with Aspose.Cells for .NET
// Description: C# sample that builds a new workbook, populates three separate tables for task completion, budget allocation, and weekly hours, then adds a column chart, a pie chart, and a line chart to distinct cell ranges on the same sheet. Each chart is linked to its own data range, given a custom title, legend settings, and saved as an XLSX file.
// Keywords: Aspose.Cells multiple charts | C# column chart Aspose.Cells | C# pie chart Aspose.Cells | C# line chart Aspose.Cells | set chart data range .NET | position chart on worksheet | Excel automation Aspose.Cells | project metrics dashboard C# | Aspose.Cells chart legend | save workbook as XLSX
// Common Searches: how to add several chart types to one worksheet using Aspose.Cells | Aspose.Cells place multiple charts in specific cell ranges | set individual data sources for each chart in Aspose.Cells .NET | C# create column, pie, and line charts on the same sheet | Aspose.Cells example for project status dashboard
// Developer Intent: Add three different chart types to one worksheet, each with its own data range and layout.
// Use Cases: Generate a project status report that visualizes task completion, budget distribution, and hours logged on a single sheet. | Build a financial dashboard with side‑by‑side charts for quick comparison of key metrics. | Automate weekly timesheet reporting by embedding a line chart for hours alongside other metric charts.
// AI Prompts: Write C# code with Aspose.Cells to insert a column chart, a pie chart, and a line chart on the same worksheet, each using a separate data range and custom title. | Show how to position three charts at different cell ranges and configure legend visibility for each chart using Aspose.Cells for .NET. | Provide an Aspose.Cells example that creates a dashboard sheet with multiple chart types and saves the workbook as XLSX.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace MultipleChartsDemo
{
    // C# sample that builds a new workbook, populates three separate tables for task completion, budget allocation, and weekly hours, then adds a column chart, a pie chart, and a line chart to distinct cell ranges on the same sheet. Each chart is linked to its own data range, given a custom title, legend settings, and saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProjectMetrics";

            // ------------------------------------------------------------
            // Populate sample data for three different metrics
            // ------------------------------------------------------------
            // Metric 1: Task Completion (%)
            sheet.Cells["A1"].PutValue("Phase");
            sheet.Cells["B1"].PutValue("Completion");
            sheet.Cells["A2"].PutValue("Planning");
            sheet.Cells["A3"].PutValue("Design");
            sheet.Cells["A4"].PutValue("Development");
            sheet.Cells["A5"].PutValue("Testing");
            sheet.Cells["B2"].PutValue(80);
            sheet.Cells["B3"].PutValue(60);
            sheet.Cells["B4"].PutValue(40);
            sheet.Cells["B5"].PutValue(20);

            // Metric 2: Budget (in thousands)
            sheet.Cells["D1"].PutValue("Category");
            sheet.Cells["E1"].PutValue("Budget");
            sheet.Cells["D2"].PutValue("Hardware");
            sheet.Cells["D3"].PutValue("Software");
            sheet.Cells["D4"].PutValue("Licensing");
            sheet.Cells["D5"].PutValue("Training");
            sheet.Cells["E2"].PutValue(120);
            sheet.Cells["E3"].PutValue(200);
            sheet.Cells["E4"].PutValue(80);
            sheet.Cells["E5"].PutValue(50);

            // Metric 3: Hours Logged
            sheet.Cells["G1"].PutValue("Week");
            sheet.Cells["H1"].PutValue("Hours");
            for (int i = 2; i <= 9; i++)
            {
                sheet.Cells[$"G{i}"].PutValue($"Week {i - 1}");
                sheet.Cells[$"H{i}"].PutValue(40 + (i - 2) * 5); // sample increasing hours
            }

            // ------------------------------------------------------------
            // Add a Column chart for Task Completion
            // ------------------------------------------------------------
            int colChartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 7);
            Chart colChart = sheet.Charts[colChartIdx];
            colChart.SetChartDataRange("A1:B5", true); // include headers
            colChart.Title.Text = "Task Completion by Phase";
            colChart.Title.TextHorizontalAlignment = TextAlignmentType.Center;
            colChart.ShowLegend = false;

            // ------------------------------------------------------------
            // Add a Pie chart for Budget distribution
            // ------------------------------------------------------------
            int pieChartIdx = sheet.Charts.Add(ChartType.Pie, 7, 8, 22, 15);
            Chart pieChart = sheet.Charts[pieChartIdx];
            // Use only the numeric values; category data is taken from column D
            pieChart.NSeries.Add("E2:E5", true);
            pieChart.NSeries.CategoryData = "D2:D5";
            pieChart.Title.Text = "Budget Allocation";
            pieChart.Title.TextHorizontalAlignment = TextAlignmentType.Center;
            pieChart.ShowLegend = true;

            // ------------------------------------------------------------
            // Add a Line chart for Hours Logged over weeks
            // ------------------------------------------------------------
            int lineChartIdx = sheet.Charts.Add(ChartType.Line, 23, 0, 38, 7);
            Chart lineChart = sheet.Charts[lineChartIdx];
            lineChart.SetChartDataRange("G1:H9", true);
            lineChart.Title.Text = "Hours Logged per Week";
            lineChart.Title.TextHorizontalAlignment = TextAlignmentType.Center;
            lineChart.ShowLegend = false;

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("ProjectMetricsCharts.xlsx", SaveFormat.Xlsx);
        }
    }
}
