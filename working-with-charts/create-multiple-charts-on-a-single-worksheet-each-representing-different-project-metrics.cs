// Title: Create Column, Pie, and Line Charts on a Single Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a new workbook, fill three metric tables (task completion, budget, risk), place a column chart, a pie chart, and a line chart on the same sheet using Aspose.Cells, and save the result as ProjectMetricsCharts.xlsx.
// Keywords: Aspose.Cells multiple charts C# | add column chart Aspose.Cells | pie chart Aspose.Cells example | line chart on same worksheet Aspose.Cells | chart positioning Aspose.Cells | set chart series data Aspose.Cells | save workbook with charts Aspose.Cells | C# Excel chart automation
// Common Searches: how to add several chart types to one worksheet using Aspose.Cells | Aspose.Cells chart location and size C# | assign category and value ranges for Aspose.Cells charts | Aspose.Cells create column, pie, line charts together | C# generate Excel file with multiple charts Aspose
// Developer Intent: Generate a single Excel worksheet that contains separate column, pie, and line charts for different project metrics and export it as an XLSX file.
// Use Cases: Show task‑completion percentages per phase with a column chart positioned in rows 7‑22, columns A‑G. | Visualize budget allocation across phases with a pie chart placed in rows 7‑22, columns H‑O. | Track risk‑level trends over phases using a line chart located in rows 24‑38, columns A‑G.
// AI Prompts: Write C# code with Aspose.Cells to add a stacked bar chart to an existing worksheet, using categories in C2:C6 and values in D2:D6, and place it at a custom position. | Explain how to customize titles, legends, and data labels for multiple charts on the same worksheet in Aspose.Cells. | Provide an example of exporting a workbook that contains several different chart types to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace MultipleChartsDemo
{
    // Demonstrates how to build a new workbook, fill three metric tables (task completion, budget, risk), place a column chart, a pie chart, and a line chart on the same sheet using Aspose.Cells, and save the result as ProjectMetricsCharts.xlsx.
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
            sheet.Cells["A4"].PutValue("Implementation");
            sheet.Cells["A5"].PutValue("Testing");
            sheet.Cells["B2"].PutValue(80);
            sheet.Cells["B3"].PutValue(60);
            sheet.Cells["B4"].PutValue(40);
            sheet.Cells["B5"].PutValue(20);

            // Metric 2: Budget (in thousands)
            sheet.Cells["D1"].PutValue("Phase");
            sheet.Cells["E1"].PutValue("Budget");
            sheet.Cells["D2"].PutValue("Planning");
            sheet.Cells["D3"].PutValue("Design");
            sheet.Cells["D4"].PutValue("Implementation");
            sheet.Cells["D5"].PutValue("Testing");
            sheet.Cells["E2"].PutValue(120);
            sheet.Cells["E3"].PutValue(250);
            sheet.Cells["E4"].PutValue(400);
            sheet.Cells["E5"].PutValue(150);

            // Metric 3: Risk Level (1-5)
            sheet.Cells["G1"].PutValue("Phase");
            sheet.Cells["H1"].PutValue("Risk");
            sheet.Cells["G2"].PutValue("Planning");
            sheet.Cells["G3"].PutValue("Design");
            sheet.Cells["G4"].PutValue("Implementation");
            sheet.Cells["G5"].PutValue("Testing");
            sheet.Cells["H2"].PutValue(2);
            sheet.Cells["H3"].PutValue(3);
            sheet.Cells["H4"].PutValue(4);
            sheet.Cells["H5"].PutValue(2);

            // ------------------------------------------------------------
            // Add a Column chart for Task Completion
            // ------------------------------------------------------------
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int colChartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 7);
            Chart colChart = sheet.Charts[colChartIdx];
            colChart.Title.Text = "Task Completion by Phase";
            // Data range includes category (A2:A5) and values (B2:B5)
            colChart.NSeries.Add("B2:B5", true);
            colChart.NSeries.CategoryData = "A2:A5";

            // ------------------------------------------------------------
            // Add a Pie chart for Budget distribution
            // ------------------------------------------------------------
            int pieChartIdx = sheet.Charts.Add(ChartType.Pie, 7, 8, 22, 15);
            Chart pieChart = sheet.Charts[pieChartIdx];
            pieChart.Title.Text = "Budget Allocation";
            // Pie charts use a single series; categories are in D2:D5, values in E2:E5
            pieChart.NSeries.Add("E2:E5", true);
            pieChart.NSeries.CategoryData = "D2:D5";

            // ------------------------------------------------------------
            // Add a Line chart for Risk Level over phases
            // ------------------------------------------------------------
            int lineChartIdx = sheet.Charts.Add(ChartType.Line, 24, 0, 38, 7);
            Chart lineChart = sheet.Charts[lineChartIdx];
            lineChart.Title.Text = "Risk Trend";
            lineChart.NSeries.Add("H2:H5", true);
            lineChart.NSeries.CategoryData = "G2:G5";

            // ------------------------------------------------------------
            // Save the workbook to an XLSX file
            // ------------------------------------------------------------
            workbook.Save("ProjectMetricsCharts.xlsx", SaveFormat.Xlsx);
        }
    }
}
