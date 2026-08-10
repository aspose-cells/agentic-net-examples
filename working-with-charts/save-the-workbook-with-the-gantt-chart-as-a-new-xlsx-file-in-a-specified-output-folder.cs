// Title: Save a Gantt‑Chart Workbook as XLSX to a Specified Folder with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, adds task rows, computes duration with formulas, builds a stacked‑bar Gantt chart, hides the start series, sets a title, and writes the file to a user‑defined output directory as GanttChart.xlsx.
// Keywords: Aspose.Cells C# | save workbook to folder | export Gantt chart XLSX | stacked bar chart Excel | C# generate Excel file | custom output directory .NET | project schedule Excel export | Aspose.Cells save example
// Common Searches: how to save an Aspose.Cells workbook to a custom folder | C# generate Gantt chart Excel file with Aspose.Cells | Aspose.Cells export stacked bar chart as XLSX | save Excel file to user‑specified directory .NET | create and store Gantt chart workbook programmatically
// Developer Intent: Generate a Gantt‑chart workbook and persist it as an XLSX file in a folder supplied at runtime.
// Use Cases: Automated nightly reporting that writes a project‑schedule Excel file to a shared network drive. | Web API endpoint that creates a Gantt chart on‑the‑fly, saves it to a temporary folder, and returns the file path for download. | Desktop utility that lets end‑users choose a destination folder for the exported Gantt‑chart workbook.
// AI Prompts: Modify the example so the start series is completely transparent instead of only hiding its border. | Add data labels to the duration series before saving the workbook. | Allow the caller to specify both the output folder and a custom file name while preserving the save logic.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a new Workbook, adds task rows, computes duration with formulas, builds a stacked‑bar Gantt chart, hides the start series, sets a title, and writes the file to a user‑defined output directory as GanttChart.xlsx.
public class GanttChartSaver
{
    public static void Run(string outputFolder)
    {
        try
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Create a new workbook and get the first worksheet.
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Header row.
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Finish");
                sheet.Cells["D1"].PutValue("Duration");

                // Sample data.
                sheet.Cells["A2"].PutValue("Planning");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 5));

                sheet.Cells["A3"].PutValue("Design");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 6));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 1, 12));

                sheet.Cells["A4"].PutValue("Implementation");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 13));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 1, 25));

                // Calculate duration (Finish - Start) in days.
                sheet.Cells["D2"].Formula = "C2-B2";
                sheet.Cells["D3"].Formula = "C3-B3";
                sheet.Cells["D4"].Formula = "C4-B4";

                // Add a stacked bar chart to simulate a Gantt chart.
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 8);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Series: Start (invisible) and Duration (visible).
                ganttChart.NSeries.Add("B2:B4", true); // Start dates
                ganttChart.NSeries.Add("D2:D4", true); // Duration
                ganttChart.NSeries.CategoryData = "A2:A4";

                // Hide the start series by making its border invisible.
                if (ganttChart.NSeries.Count > 0)
                {
                    Series startSeries = ganttChart.NSeries[0];
                    startSeries.Border.IsVisible = false; // Hide border if any
                }

                // Optional: set chart title.
                ganttChart.Title.Text = "Project Schedule";

                // Build the full file path.
                string filePath = Path.Combine(outputFolder, "GanttChart.xlsx");

                // Save the workbook.
                workbook.Save(filePath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook with Gantt chart saved to: {filePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating Gantt chart workbook: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Determine output folder (use current directory if not provided).
        string outputFolder = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        GanttChartSaver.Run(outputFolder);
    }
}
