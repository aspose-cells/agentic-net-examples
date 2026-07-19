// Title: Save a Gantt‑Chart workbook to a custom folder using Aspose.Cells for .NET (C#)
// Description: This example creates a new Workbook, inserts task data, builds a stacked‑bar chart that mimics a Gantt chart by making the start series transparent, ensures the target directory exists, and saves the file as an XLSX workbook to a user‑specified output folder.
// Keywords: Aspose.Cells C# | save workbook to folder | Gantt chart Excel | stacked bar chart Aspose | export XLSX .NET | create project schedule Excel | output directory Aspose.Cells | hide start series Gantt
// Common Searches: Aspose.Cells save workbook to custom directory | C# generate Gantt chart Excel file | how to export stacked bar chart as Gantt chart using Aspose.Cells | create and save Excel workbook with Gantt chart .NET | ensure output folder exists before saving Aspose.Cells workbook
// Developer Intent: Generate an Excel workbook that contains a Gantt‑style chart and write the file to a specified folder.
// Use Cases: Automate weekly project schedule reports with Gantt charts and store them in a shared network folder. | Produce per‑project Gantt‑chart workbooks during a build process and publish them to the CI/CD artifacts directory. | Create a desktop utility that lets users select an output path and saves a ready‑made Gantt chart for project planning.
// AI Prompts: Write C# code with Aspose.Cells that builds a Gantt chart from task data and saves the workbook to a given path. | Explain how to make the start series of a stacked bar chart transparent to simulate a Gantt chart in Aspose.Cells. | Add parameters for output folder and file name, include folder‑creation logic, and handle save‑time exceptions.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a new Workbook, inserts task data, builds a stacked‑bar chart that mimics a Gantt chart by making the start series transparent, ensures the target directory exists, and saves the file as an XLSX workbook to a user‑specified output folder.
class SaveGanttChart
{
    static void Main()
    {
        try
        {
            // Define the output folder and ensure it exists
            string outputFolder = @"C:\Output\Gantt";
            Directory.CreateDirectory(outputFolder);

            // Full path for the new XLSX file
            string outputPath = Path.Combine(outputFolder, "GanttChart.xlsx");

            // Create a new workbook
            using (Workbook workbook = new Workbook())
            {
                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the Gantt chart
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(10);

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 11));
                sheet.Cells["C3"].PutValue(20);

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 1));
                sheet.Cells["C4"].PutValue(5);

                // Add a stacked bar chart to simulate a Gantt chart
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (Start and Duration)
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the "Start" series so only the duration appears as a Gantt bar
                chart.NSeries[0].Area.ForegroundColor = Color.Transparent;

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook with Gantt chart saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
