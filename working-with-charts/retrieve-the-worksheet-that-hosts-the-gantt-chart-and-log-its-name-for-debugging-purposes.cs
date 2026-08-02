// Title: Get the Worksheet Hosting a Gantt Chart with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a "ProjectPlan" sheet, fills it with task dates, inserts a stacked‑bar chart that mimics a Gantt chart, then uses the Chart.Worksheet property to retrieve the parent worksheet and writes its name to the console before saving the file.
// Keywords: Aspose.Cells Chart.Worksheet | retrieve chart worksheet C# | Gantt chart Aspose.Cells | log worksheet name .NET | debug Aspose.Cells chart | C# Aspose.Cells example
// Common Searches: Aspose.Cells get worksheet from chart C# | Chart.Worksheet property example .NET | How to find the sheet that contains a Gantt chart using Aspose.Cells | Log chart host sheet name Aspose.Cells | Retrieve parent worksheet of a stacked bar chart C#
// Developer Intent: Identify the worksheet that contains a specific chart and output its name for debugging or validation.
// Use Cases: Verify that a generated Gantt chart is placed on the correct sheet before publishing a report. | Log worksheet names of multiple charts when automating workbook creation for audit trails. | Programmatically move or delete a chart only after confirming its parent worksheet.
// AI Prompts: Show how to use Aspose.Cells Chart.Worksheet to obtain the host worksheet and print its name in C#. | Provide a loop that iterates through all charts in a workbook and displays each chart’s worksheet name using Aspose.Cells for .NET. | Explain how to handle a null Chart.Worksheet reference safely when retrieving the parent sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartWorksheetDemo
{
    // This example creates a workbook, adds a "ProjectPlan" sheet, fills it with task dates, inserts a stacked‑bar chart that mimics a Gantt chart, then uses the Chart.Worksheet property to retrieve the parent worksheet and writes its name to the console before saving the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet where the Gantt chart will be placed
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "ProjectPlan";

                // Populate sample data for the Gantt chart
                // Columns: Task, Start Date, End Date
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("End");

                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 15));

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 16));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 2, 28));

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 3, 15));

                // Add a stacked bar chart to simulate a Gantt chart
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                // Category (Task) data
                ganttChart.NSeries.CategoryData = "A2:A4";
                // Series data: Start and End dates
                ganttChart.NSeries.Add("B2:C4", true);

                // Retrieve the worksheet that hosts the chart using the Worksheet property
                Worksheet hostWorksheet = ganttChart.Worksheet;

                // Log the worksheet name for debugging purposes
                Console.WriteLine("Gantt chart is hosted on worksheet: " + hostWorksheet.Name);

                // Save the workbook (lifecycle rule: save)
                workbook.Save("GanttChartDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
