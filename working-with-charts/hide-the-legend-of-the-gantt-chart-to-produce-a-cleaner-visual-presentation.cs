using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class HideGanttLegend
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample data for a Gantt‑like chart
            // Column A: Task names, Column B: Start dates, Column C: Duration (in days)
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Start");
            sheet.Cells["C1"].PutValue("Duration");

            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
            sheet.Cells["C3"].PutValue(7);

            // Format the start date column as a date (optional, improves appearance)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            // Apply the style to the range B2:B3 (Cell does not have ApplyStyle)
            sheet.Cells.CreateRange("B2:B3").ApplyStyle(dateStyle, flag);

            // Add a stacked bar chart (used here to emulate a Gantt chart)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart:
            //   Series data (Start and Duration) -> B2:C3
            //   Category (Task names)          -> A2:A3
            chart.NSeries.Add("B2:C3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Hide the legend to produce a cleaner visual presentation
            chart.ShowLegend = false;

            // Define output file name
            string outputPath = "GanttChart_NoLegend.xlsx";

            // Save the workbook (no template file to check)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}