// Title: Generate several different charts on one worksheet to display project metrics with Aspose.Cells in C#
// Common Searches: how to add multiple charts to a single worksheet using Aspose.Cells C# | Aspose.Cells C# create column, pie and line charts on the same sheet | sample code for generating project metric charts in Excel with Aspose.Cells | position several charts on one worksheet Aspose.Cells example | add task completion column chart and budget allocation pie chart with Aspose.Cells
// Tags: add multiple charts to worksheet Aspose.Cells C# | column chart task completion Aspose.Cells | pie chart budget allocation Aspose.Cells | line chart hours spent Aspose.Cells | chart positioning on Excel sheet Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProjectMetricsCharts
{
    // The example creates a new workbook, adds a worksheet named "Metrics", fills three data tables (task completion, budget allocation, hours spent), and inserts a column chart, a pie chart, and a line chart—each placed separately on the same sheet—before saving the file as ProjectMetrics.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Metrics";

                // Populate sample data for three different metrics

                // Metric 1: Task Completion (%)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Completion %");
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(80);
                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(55);
                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(30);

                // Metric 2: Budget Allocation (USD)
                sheet.Cells["D1"].PutValue("Phase");
                sheet.Cells["E1"].PutValue("Budget");
                sheet.Cells["D2"].PutValue("Design");
                sheet.Cells["E2"].PutValue(15000);
                sheet.Cells["D3"].PutValue("Development");
                sheet.Cells["E3"].PutValue(35000);
                sheet.Cells["D4"].PutValue("Testing");
                sheet.Cells["E4"].PutValue(20000);

                // Metric 3: Hours Spent
                sheet.Cells["G1"].PutValue("Phase");
                sheet.Cells["H1"].PutValue("Hours");
                sheet.Cells["G2"].PutValue("Design");
                sheet.Cells["H2"].PutValue(120);
                sheet.Cells["G3"].PutValue("Development");
                sheet.Cells["H3"].PutValue(300);
                sheet.Cells["G4"].PutValue("Testing");
                sheet.Cells["H4"].PutValue(180);

                // -------------------------
                // Create Chart 1: Column Chart for Task Completion
                // -------------------------
                int chartIndex1 = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 7);
                Chart chart1 = sheet.Charts[chartIndex1];
                chart1.Title.Text = "Task Completion Percentage";
                chart1.NSeries.Add("B2:B4", true);
                // Category (X‑axis) data is taken from the first column (A2:A4) by default.

                // -------------------------
                // Create Chart 2: Pie Chart for Budget Allocation
                // -------------------------
                int chartIndex2 = sheet.Charts.Add(ChartType.Pie, 6, 8, 20, 15);
                Chart chart2 = sheet.Charts[chartIndex2];
                chart2.Title.Text = "Budget Allocation";
                chart2.NSeries.Add("E2:E4", true);
                // Category data defaults to D2:D4.

                // -------------------------
                // Create Chart 3: Line Chart for Hours Spent
                // -------------------------
                int chartIndex3 = sheet.Charts.Add(ChartType.Line, 22, 0, 36, 7);
                Chart chart3 = sheet.Charts[chartIndex3];
                chart3.Title.Text = "Hours Spent per Phase";
                chart3.NSeries.Add("H2:H4", true);
                // Category data defaults to G2:G4.

                // Save the workbook to a file
                string outputPath = "ProjectMetrics.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
