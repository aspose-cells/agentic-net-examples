// Title: Generate column, line, and pie charts on the same worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, fills it with monthly project data, and adds a column chart for tasks completed, a line chart for budget used, and a pie chart for hours spent, each positioned in separate ranges on the same sheet. | Demonstrate how to set chart titles, series names, category labels, and then save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: how to add multiple different chart types to a single worksheet with Aspose.Cells C# | Aspose.Cells place column, line and pie charts on the same sheet with custom positions | C# example for creating project metrics dashboard using Aspose.Cells charts | set category axis labels for a column chart in Aspose.Cells .NET | save workbook containing several charts to Excel file using Aspose.Cells
// Tags: create multiple chart types Aspose.Cells C# | position charts on worksheet Aspose.Cells | define chart series range Aspose.Cells | add column chart with category labels Aspose.Cells | export workbook with charts to XLSX Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Shows how to build a workbook, populate it with sample project metric data, and add three distinct charts (column, line, pie) with specific positions, titles, and series on a single worksheet, then save the file as ProjectMetrics.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a meaningful name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProjectMetrics";

            // -------------------------------------------------
            // Populate sample data for three different metrics
            // -------------------------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Tasks Completed");
            sheet.Cells["C1"].PutValue("Budget Used");
            sheet.Cells["D1"].PutValue("Hours Spent");

            // Sample data
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            int[] tasks = { 20, 35, 30, 45, 50 };
            double[] budget = { 5000, 7000, 6500, 8000, 9000 };
            int[] hours = { 120, 150, 130, 170, 200 };

            for (int i = 0; i < months.Length; i++)
            {
                int row = i + 1; // Data starts at row 2 (index 1)
                sheet.Cells[row, 0].PutValue(months[i]);   // Column A
                sheet.Cells[row, 1].PutValue(tasks[i]);    // Column B
                sheet.Cells[row, 2].PutValue(budget[i]);   // Column C
                sheet.Cells[row, 3].PutValue(hours[i]);    // Column D
            }

            // -------------------------------------------------
            // Create a Column chart for "Tasks Completed"
            // -------------------------------------------------
            // Parameters: ChartType, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart1 = sheet.Charts[chartIdx1];
            chart1.Title.Text = "Tasks Completed per Month";
            // Add series: data range B2:B6, use categories from column A
            chart1.NSeries.Add("B2:B6", true);
            chart1.NSeries[0].Name = "Tasks Completed";

            // -------------------------------------------------
            // Create a Line chart for "Budget Used"
            // -------------------------------------------------
            int chartIdx2 = sheet.Charts.Add(ChartType.Line, 5, 8, 20, 15);
            Chart chart2 = sheet.Charts[chartIdx2];
            chart2.Title.Text = "Budget Used per Month";
            chart2.NSeries.Add("C2:C6", true);
            chart2.NSeries[0].Name = "Budget Used";

            // -------------------------------------------------
            // Create a Pie chart for "Hours Spent"
            // -------------------------------------------------
            int chartIdx3 = sheet.Charts.Add(ChartType.Pie, 22, 0, 35, 7);
            Chart chart3 = sheet.Charts[chartIdx3];
            chart3.Title.Text = "Hours Spent Distribution";
            chart3.NSeries.Add("D2:D6", true);
            chart3.NSeries[0].Name = "Hours Spent";

            // -------------------------------------------------
            // Save the workbook to a file
            // -------------------------------------------------
            string outputPath = "ProjectMetrics.xlsx";
            // Ensure the directory exists (useful if a path is provided)
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
