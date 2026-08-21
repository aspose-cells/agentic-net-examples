// Title: Retrieve the worksheet that hosts a Gantt chart and log its name with Aspose.Cells (C#)
// Description: Creates a workbook, adds a stacked‑bar chart to simulate a Gantt chart, accesses the chart's Worksheet property, prints the sheet name to the console, and saves the file. Demonstrates how to identify the parent worksheet of any chart in Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart worksheet | get chart parent sheet | Gantt chart worksheet name | Aspose.Cells retrieve chart sheet | log worksheet name Aspose.Cells
// Common Searches: how to find worksheet of a chart in Aspose.Cells | Aspose.Cells get worksheet name from chart | C# retrieve parent sheet of Gantt chart | debug chart location Aspose.Cells | Aspose.Cells chart.Worksheet property example
// Developer Intent: Obtain the worksheet that contains a specific chart and output its name for verification or further processing.
// Use Cases: Confirm that a generated Gantt chart is placed on the correct sheet before publishing a report. | Log worksheet names of multiple charts during automated workbook generation to detect placement errors. | Use the returned Worksheet object to apply additional formatting or data updates to the chart's host sheet.
// AI Prompts: Show C# code to get the parent worksheet of any chart in an Aspose.Cells workbook and print its name. | Write a script that iterates over all charts in a workbook, logs each chart's worksheet name, and saves the workbook. | Explain how to handle charts located on hidden worksheets when retrieving their Worksheet property with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a stacked‑bar chart to simulate a Gantt chart, accesses the chart's Worksheet property, prints the sheet name to the console, and saves the file. Demonstrates how to identify the parent worksheet of any chart in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a meaningful name
            Worksheet ws = workbook.Worksheets[0];
            ws.Name = "Project";

            // Populate sample data required for a Gantt chart
            ws.Cells["A1"].PutValue("Task");
            ws.Cells["B1"].PutValue("Start");
            ws.Cells["C1"].PutValue("Duration");

            ws.Cells["A2"].PutValue("Task 1");
            ws.Cells["B2"].PutValue(DateTime.Today);
            ws.Cells["C2"].PutValue(5);

            ws.Cells["A3"].PutValue("Task 2");
            ws.Cells["B3"].PutValue(DateTime.Today.AddDays(2));
            ws.Cells["C3"].PutValue(3);

            // Add a stacked bar chart (used to emulate a Gantt chart)
            int chartIndex = ws.Charts.Add(ChartType.BarStacked, 5, 0, 15, 5);
            Chart ganttChart = ws.Charts[chartIndex];

            // First series: Start dates (will be made invisible)
            ganttChart.NSeries.Add("B2:B3", true);
            // Second series: Duration values
            ganttChart.NSeries.Add("C2:C3", true);

            // Category (task names)
            ganttChart.NSeries.CategoryData = "A2:A3";

            // Hide the start series to create the Gantt effect
            ganttChart.NSeries[0].Area.ForegroundColor = System.Drawing.Color.Transparent;
            ganttChart.NSeries[0].Border.IsVisible = false;

            // Retrieve the worksheet that hosts the Gantt chart
            Worksheet chartWorksheet = ganttChart.Worksheet;

            // Log the worksheet name for debugging purposes
            Console.WriteLine("Gantt chart is on worksheet: " + chartWorksheet.Name);

            // Save the workbook
            workbook.Save("GanttChartDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
