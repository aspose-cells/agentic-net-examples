// Title: Generate column, line, and pie charts on the same worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart for tasks completed per month, a line chart for budget planned vs spent, and a pie chart for issue type distribution on one worksheet using Aspose.Cells in C#. | Specify start and end cell coordinates to position each chart when adding them to the worksheet with Aspose.Cells. | Add a stacked bar chart for task data while keeping the existing column, line, and pie charts on the same sheet using Aspose.Cells.
// Common Searches: aspnet how to place multiple chart types on a single Excel sheet with Aspose.Cells | c# Aspose.Cells add column chart and line chart together on one worksheet | example code for creating a pie chart and line chart on the same sheet using Aspose.Cells for .NET | Aspose.Cells set chart location using cell range in C# | generate project metrics dashboard with several charts in one worksheet Aspose.Cells
// Tags: Aspose.Cells create column chart C# | Aspose.Cells line chart budget comparison | Aspose.Cells pie chart issue distribution | Aspose.Cells multiple charts same worksheet | Aspose.Cells set chart position by cell range

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Demonstrates building a workbook, populating three metric tables, and adding a column chart, a line chart, and a pie chart to the same worksheet with Aspose.Cells for .NET, then saving the file as ProjectMetrics.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProjectMetrics";

            // -------------------- Fill sample data --------------------
            // Metric 1: Tasks Completed per Month
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Tasks Completed");
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            int[] tasks = { 20, 35, 30, 45, 50 };
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(tasks[i]); // Column B
            }

            // Metric 2: Budget Planned vs Spent
            sheet.Cells["D1"].PutValue("Month");
            sheet.Cells["E1"].PutValue("Planned");
            sheet.Cells["F1"].PutValue("Spent");
            int[] planned = { 100, 120, 130, 150, 170 };
            int[] spent = { 90, 110, 140, 160, 180 };
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 3].PutValue(months[i]); // Column D
                sheet.Cells[i + 1, 4].PutValue(planned[i]); // Column E
                sheet.Cells[i + 1, 5].PutValue(spent[i]); // Column F
            }

            // Metric 3: Issue Types Distribution
            sheet.Cells["H1"].PutValue("Issue Type");
            sheet.Cells["I1"].PutValue("Count");
            string[] issueTypes = { "Bug", "Feature", "Improvement" };
            int[] counts = { 15, 8, 12 };
            for (int i = 0; i < issueTypes.Length; i++)
            {
                sheet.Cells[i + 1, 7].PutValue(issueTypes[i]); // Column H
                sheet.Cells[i + 1, 8].PutValue(counts[i]); // Column I
            }

            // -------------------- Create charts --------------------
            // Column chart for Tasks Completed
            int chartIndex1 = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 7);
            Chart chart1 = sheet.Charts[chartIndex1];
            chart1.Title.Text = "Tasks Completed per Month";
            chart1.NSeries.Add("B2:B6", true);               // Y values
            chart1.NSeries.CategoryData = "A2:A6";           // X categories

            // Line chart for Budget Planned vs Spent
            int chartIndex2 = sheet.Charts.Add(ChartType.Line, 7, 8, 22, 15);
            Chart chart2 = sheet.Charts[chartIndex2];
            chart2.Title.Text = "Budget: Planned vs Spent";
            chart2.NSeries.Add("E2:E6", true);               // Planned series
            chart2.NSeries.Add("F2:F6", true);               // Spent series
            chart2.NSeries.CategoryData = "D2:D6";           // X categories (Month)

            // Pie chart for Issue Types Distribution
            int chartIndex3 = sheet.Charts.Add(ChartType.Pie, 24, 0, 38, 7);
            Chart chart3 = sheet.Charts[chartIndex3];
            chart3.Title.Text = "Issue Types Distribution";
            chart3.NSeries.Add("I2:I4", true);               // Values
            chart3.NSeries.CategoryData = "H2:H4";           // Labels

            // -------------------- Save workbook --------------------
            string outputPath = "ProjectMetrics.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
