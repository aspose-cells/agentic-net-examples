// Title: Hide the legend in a Gantt chart created with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide legend in Gantt chart | remove chart legend from bar chart using Aspose.Cells .NET | how to delete legend in Excel chart with Aspose.Cells C# | programmatically hide Gantt chart legend Aspose.Cells | set chart.Legend.IsDeleted true in C# Aspose.Cells
// Tags: Aspose.Cells hide chart legend | Gantt chart legend removal Aspose.Cells | Bar chart legend delete .NET | Excel chart legend suppression C# | chart.Legend.IsDeleted usage

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace GanttChartHideLegend
{
    // The example creates a workbook, adds task, start date, and duration data, formats the start column as dates, builds a bar chart that functions as a Gantt chart, makes the start series invisible, removes bar gaps, and hides the chart legend by setting chart.Legend.IsDeleted to true before saving the file as GanttChart_NoLegend.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Prepare sample data for a simple Gantt chart
                // Columns: Task Name | Start Date (as number) | Duration (days)
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                sheet.Cells["A2"].PutValue("Planning");
                sheet.Cells["A3"].PutValue("Design");
                sheet.Cells["A4"].PutValue("Development");
                sheet.Cells["A5"].PutValue("Testing");

                // Use Excel serial dates (e.g., 44561 = 2021-12-01)
                sheet.Cells["B2"].PutValue(44561); // 2021-12-01
                sheet.Cells["B3"].PutValue(44566); // 2021-12-06
                sheet.Cells["B4"].PutValue(44571); // 2021-12-11
                sheet.Cells["B5"].PutValue(44581); // 2021-12-21

                sheet.Cells["C2"].PutValue(5);
                sheet.Cells["C3"].PutValue(7);
                sheet.Cells["C4"].PutValue(10);
                sheet.Cells["C5"].PutValue(4);

                // Format the Start column as a date for visual clarity
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Built‑in date format (mm-dd-yy)
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                AsposeRange dateRange = sheet.Cells.CreateRange("B2:B5");
                dateRange.ApplyStyle(dateStyle, flag);

                // ------------------------------------------------------------
                // Add a bar chart that will act as a Gantt chart
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // First series: Start (invisible)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Start";
                // Make the start series invisible by removing its fill
                chart.NSeries[0].Area.FillFormat.Pattern = FillPattern.None;

                // Second series: Duration (visible)
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Duration";

                // Set category (task names)
                chart.NSeries.CategoryData = "A2:A5";

                // Remove gaps between bars for a tighter Gantt look
                chart.GapWidth = 0;

                // ------------------------------------------------------------
                // Hide the legend to produce a cleaner visual presentation
                // ------------------------------------------------------------
                chart.Legend.IsDeleted = true;

                // Optional: Adjust chart title
                chart.Title.Text = "Project Schedule (Gantt)";

                // Save the workbook
                string outputPath = "GanttChart_NoLegend.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'. Legend is hidden.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the Gantt chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
