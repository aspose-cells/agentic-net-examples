// Title: Save a Gantt chart workbook as an XLSX file to a custom output folder using Aspose.Cells for .NET (C#)
// Common Searches: how to save an Aspose.Cells workbook as an xlsx file in a specific directory using C# | Aspose.Cells create Gantt chart and export to Excel file in custom folder | C# code to write an Aspose.Cells workbook with a Gantt chart to a user-defined output path | example of saving a Gantt chart workbook to a folder with Aspose.Cells .NET | Aspose.Cells .NET save Excel file to C:\OutputFolder
// Tags: Aspose.Cells workbook save custom path | Aspose.Cells export Gantt chart to XLSX | Aspose.Cells bar chart Gantt representation | C# Aspose.Cells write Excel to folder | Aspose.Cells chart series hide start dates | Aspose.Cells convert DateTime to OADate

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartExample
{
    // The example creates a new workbook, adds task, start date, and duration data, converts the dates to Excel serial numbers, builds a bar chart that mimics a Gantt chart (hiding the start‑date series and value axis), and then saves the workbook as "GanttChart.xlsx" in a specified output directory.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Specify the output folder where the XLSX file will be saved
                string outputFolder = @"C:\OutputFolder";

                // Ensure the output directory exists
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "GanttData";

                // Sample data for Gantt chart
                // Columns: Task, Start Date, Duration (days)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start Date");
                sheet.Cells["C1"].PutValue("Duration");

                sheet.Cells["A2"].PutValue("Planning");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(5);

                sheet.Cells["A3"].PutValue("Design");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 6));
                sheet.Cells["C3"].PutValue(8);

                sheet.Cells["A4"].PutValue("Development");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 14));
                sheet.Cells["C4"].PutValue(15);

                sheet.Cells["A5"].PutValue("Testing");
                sheet.Cells["B5"].PutValue(new DateTime(2023, 1, 29));
                sheet.Cells["C5"].PutValue(7);

                // Convert start dates to numeric values (Excel date serial numbers)
                for (int row = 2; row <= 5; row++)
                {
                    DateTime start = sheet.Cells[row, 1].DateTimeValue;
                    sheet.Cells[row, 1].PutValue(start.ToOADate());
                }

                // Add a bar chart to simulate a Gantt chart
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Project Schedule";

                // Series 1: Start Date (invisible)
                int seriesIndex = chart.NSeries.Add("B2:B5", true);
                chart.NSeries[seriesIndex].Name = "Start";
                // Hide the first series by making its fill transparent
                chart.NSeries[seriesIndex].Area.ForegroundColor = Color.Transparent;

                // Series 2: Duration (visible)
                seriesIndex = chart.NSeries.Add("C2:C5", true);
                chart.NSeries[seriesIndex].Name = "Duration";

                // Set category axis labels (tasks)
                chart.NSeries.CategoryData = "A2:A5";

                // Hide value axis for a cleaner Gantt look
                chart.ValueAxis.IsVisible = false;

                // Save the workbook as a new XLSX file in the specified output folder
                string outputPath = Path.Combine(outputFolder, "GanttChart.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
