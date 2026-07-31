// Title: Create a Gantt chart in an existing XLSX workbook with Aspose.Cells for .NET
// Description: Load an XLSX file, add task data, insert a stacked bar chart, hide the start series, set overlap and GapWidth to zero, and save the workbook as a Gantt chart using Aspose.Cells in C#.
// Keywords: Aspose.Cells | Gantt chart | stacked bar chart | C# | ASP.NET | Excel chart overlap | GapWidth | project timeline chart | add chart to workbook | Aspose.Cells .NET example
// Common Searches: Aspose.Cells create Gantt chart from stacked bar | C# add Gantt chart to existing Excel file | set chart overlap and gap width Aspose.Cells | how to hide series in Aspose.Cells chart | generate project timeline Excel with Aspose.Cells
// Developer Intent: Generate a Gantt chart inside an existing XLSX workbook by adding a stacked bar chart and configuring series visibility, overlap, and gaps with Aspose.Cells for .NET.
// Use Cases: Automated project status reports that embed Gantt charts directly into Excel files. | Dynamic creation of timeline visualizations for task management dashboards in a C# application. | Reusable routine that converts raw start‑date and duration data into a professional Gantt view.
// AI Prompts: Show C# code that loads an XLSX workbook, adds a stacked bar chart, hides the start series, and sets Overlap and GapWidth to produce a Gantt chart with Aspose.Cells. | Explain how to configure series properties in Aspose.Cells to turn a stacked bar chart into a Gantt chart. | Provide step‑by‑step instructions for creating a project Gantt chart in Excel using Aspose.Cells for .NET, including data preparation and chart saving.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLSX file, add task data, insert a stacked bar chart, hide the start series, set overlap and GapWidth to zero, and save the workbook as a Gantt chart using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // Workbook(string) constructor

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Prepare sample data for a Gantt chart (if not already present)
        // Columns: A - Task name, B - Start (numeric), C - Duration (numeric)
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Start");
        sheet.Cells["C1"].PutValue("Duration");

        sheet.Cells["A2"].PutValue("Task 1");
        sheet.Cells["B2"].PutValue(1);   // start
        sheet.Cells["C2"].PutValue(3);   // duration

        sheet.Cells["A3"].PutValue("Task 2");
        sheet.Cells["B3"].PutValue(2);
        sheet.Cells["C3"].PutValue(4);

        sheet.Cells["A4"].PutValue("Task 3");
        sheet.Cells["B4"].PutValue(4);
        sheet.Cells["C4"].PutValue(2);

        sheet.Cells["A5"].PutValue("Task 4");
        sheet.Cells["B5"].PutValue(5);
        sheet.Cells["C5"].PutValue(5);
        // -------------------------------------------------

        // Add a stacked bar chart (BarStacked) to the worksheet
        // Parameters: topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 10); // ChartCollection.Add(ChartType, int, int, int, int)
        Chart chart = sheet.Charts[chartIndex];

        // Add the first series (Start) – this series will be made invisible later
        chart.NSeries.Add("B2:B5", true); // series for start values

        // Add the second series (Duration) – this series represents the actual task length
        chart.NSeries.Add("C2:C5", true); // series for duration values

        // Set the category (task names) for the X‑axis
        chart.NSeries.CategoryData = "A2:A5";

        // -------------------------------------------------
        // Configure the chart to behave like a Gantt chart
        // 1. Make the "Start" series invisible by setting its gap width to 0
        // 2. Overlap the bars completely (100%) so they appear stacked
        // 3. Reduce the gap between bar clusters for a tighter look
        // -------------------------------------------------
        // Make the first series (Start) have no visible gap
        chart.NSeries[0].GapWidth = 0;          // Series.GapWidth property
        // Overlap the bars completely
        chart.NSeries[0].Overlap = 100;        // Series.Overlap property
        chart.NSeries[1].Overlap = 100;        // Ensure the second series also overlaps
        // Reduce overall gap between bar clusters
        chart.GapWidth = 0;                    // Chart.GapWidth property

        // Optional: set the chart title
        chart.Title.Text = "Project Gantt Chart";

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx); // Workbook.Save(string, SaveFormat)
    }
}
