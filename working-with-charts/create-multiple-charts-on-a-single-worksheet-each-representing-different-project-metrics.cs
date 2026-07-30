// Title: Add Column, Pie, and Line Charts to One Worksheet with Aspose.Cells for .NET (C#)
// Description: C# sample that builds a new workbook, fills it with monthly project metrics, and inserts three independent charts—a column chart for tasks completed, a pie chart for budget distribution, and a line chart for hours worked—each positioned separately on the same sheet and saved as an XLSX file using Aspose.Cells.
// Keywords: Aspose.Cells C# multiple charts | create column chart Aspose.Cells | pie chart Aspose.Cells .NET | line chart Aspose.Cells example | chart positioning Aspose.Cells | Excel workbook with several charts | Aspose.Cells chart data range | C# Excel chart automation | save XLSX with charts Aspose.Cells
// Common Searches: Aspose.Cells add multiple chart types to one sheet | C# create column, pie and line charts with Aspose.Cells | set chart location programmatically Aspose.Cells | bind chart data range Aspose.Cells C# | export Excel file containing several charts using .NET
// Developer Intent: Generate an Excel file that contains three distinct charts (column, pie, line) on a single worksheet, each linked to its own data series.
// Use Cases: Produce a monthly project dashboard that visualizes tasks, budget, and hours with the most appropriate chart type on one page. | Automate reporting pipelines that need several charts in a single Excel sheet for quick stakeholder review. | Create reusable templates where different metrics are displayed side‑by‑side for comparative analysis.
// AI Prompts: Show how to add a stacked bar chart next to existing charts on the same worksheet using Aspose.Cells. | Generate code that updates chart titles and data ranges dynamically based on user‑provided column names. | Explain how to customize legends, colors, and marker styles for multiple charts on one sheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace MultipleChartsDemo
{
    // C# sample that builds a new workbook, fills it with monthly project metrics, and inserts three independent charts—a column chart for tasks completed, a pie chart for budget distribution, and a line chart for hours worked—each positioned separately on the same sheet and saved as an XLSX file using Aspose.Cells.
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
            // Header row
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Tasks Completed");
            sheet.Cells["C1"].PutValue("Budget Spent");
            sheet.Cells["D1"].PutValue("Hours Worked");

            // Sample data (Jan to Jun)
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            int[] tasks = { 20, 35, 30, 45, 40, 50 };
            double[] budget = { 5000, 7200, 6100, 8000, 7500, 9000 };
            int[] hours = { 160, 170, 165, 180, 175, 190 };

            for (int i = 0; i < months.Length; i++)
            {
                int row = i + 2; // Data starts from row 2
                sheet.Cells[row, 0].PutValue(months[i]);          // Column A
                sheet.Cells[row, 1].PutValue(tasks[i]);           // Column B
                sheet.Cells[row, 2].PutValue(budget[i]);          // Column C
                sheet.Cells[row, 3].PutValue(hours[i]);           // Column D
            }

            // ------------------------------------------------------------
            // 1st Chart: Column chart for Tasks Completed
            // ------------------------------------------------------------
            int colChartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart colChart = sheet.Charts[colChartIdx];
            colChart.Title.Text = "Tasks Completed per Month";
            // Data range: B2:B7 (tasks) and categories A2:A7 (months)
            colChart.NSeries.Add("B2:B7", true);
            colChart.NSeries.CategoryData = "A2:A7";

            // ------------------------------------------------------------
            // 2nd Chart: Pie chart for Budget Distribution
            // ------------------------------------------------------------
            int pieChartIdx = sheet.Charts.Add(ChartType.Pie, 22, 0, 37, 8);
            Chart pieChart = sheet.Charts[pieChartIdx];
            pieChart.Title.Text = "Budget Spent Distribution";
            // Data range: C2:C7 (budget) and categories A2:A7 (months)
            pieChart.NSeries.Add("C2:C7", true);
            pieChart.NSeries.CategoryData = "A2:A7";

            // ------------------------------------------------------------
            // 3rd Chart: Line chart for Hours Worked
            // ------------------------------------------------------------
            int lineChartIdx = sheet.Charts.Add(ChartType.Line, 39, 0, 54, 8);
            Chart lineChart = sheet.Charts[lineChartIdx];
            lineChart.Title.Text = "Hours Worked per Month";
            // Data range: D2:D7 (hours) and categories A2:A7 (months)
            lineChart.NSeries.Add("D2:D7", true);
            lineChart.NSeries.CategoryData = "A2:A7";

            // ------------------------------------------------------------
            // Save the workbook to an XLSX file
            // ------------------------------------------------------------
            workbook.Save("ProjectMetricsCharts.xlsx", SaveFormat.Xlsx);
        }
    }
}
