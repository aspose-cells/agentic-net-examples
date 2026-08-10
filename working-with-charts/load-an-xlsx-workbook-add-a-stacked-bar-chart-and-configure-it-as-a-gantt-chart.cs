// Title: Create a Gantt chart in C# with Aspose.Cells by adding a stacked bar chart to an XLSX workbook
// Description: C# sample that loads (or creates) an XLSX file containing Task, Start and Duration columns, inserts a stacked bar chart, hides the Start series, sets the gap width to zero, adds a chart title, and saves the file as a Gantt chart.
// Keywords: Aspose.Cells | C# | Gantt chart | stacked bar chart | Excel chart | transparent series | gap width | project schedule | load workbook | add chart
// Common Searches: Aspose.Cells create Gantt chart C# | how to hide series in Aspose.Cells chart | stacked bar chart Gantt view Aspose | set gap width zero Aspose.Cells | generate project schedule chart with Aspose.Cells
// Developer Intent: Generate a Gantt chart from task data in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Transform a task list with start dates and durations into a visual Gantt chart for project reporting. | Automatically create a sample workbook when the source file is missing and produce a Gantt chart. | Customize chart appearance—gap width, transparent start series, title—for inclusion in dashboards or presentations.
// AI Prompts: Write C# code using Aspose.Cells to load an Excel file, add a stacked bar chart, hide the start series, and output a Gantt chart. | Show how to set GapWidth to zero and make a series transparent in an Aspose.Cells chart for a Gantt view. | Provide a complete example that creates a sample task table if the input file does not exist and then generates a Gantt chart.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttExample
{
    // C# sample that loads (or creates) an XLSX file containing Task, Start and Duration columns, inserts a stacked bar chart, hides the Start series, sets the gap width to zero, adds a chart title, and saves the file as a Gantt chart.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; if not, create a sample workbook
                if (!File.Exists(inputPath))
                {
                    var sampleWorkbook = new Workbook();
                    var sheet = sampleWorkbook.Worksheets[0];

                    // Header row
                    sheet.Cells["A1"].PutValue("Task");
                    sheet.Cells["B1"].PutValue("Start");
                    sheet.Cells["C1"].PutValue("Duration");

                    // Sample data (rows 2‑6)
                    string[] tasks = { "Task 1", "Task 2", "Task 3", "Task 4", "Task 5" };
                    int[] starts = { 0, 2, 4, 6, 8 };
                    int[] durations = { 2, 3, 1, 4, 2 };

                    for (int i = 0; i < tasks.Length; i++)
                    {
                        sheet.Cells[i + 1, 0].PutValue(tasks[i]);   // Column A
                        sheet.Cells[i + 1, 1].PutValue(starts[i]); // Column B
                        sheet.Cells[i + 1, 2].PutValue(durations[i]); // Column C
                    }

                    sampleWorkbook.Save(inputPath, SaveFormat.Xlsx);
                }

                // Load the workbook (ensure file exists before loading)
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook '{inputPath}': {loadEx.Message}");
                    return;
                }

                var worksheet = workbook.Worksheets[0];

                // Add a stacked bar chart (rows 5‑20, columns 0‑10)
                int chartIndex = worksheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
                Chart ganttChart = worksheet.Charts[chartIndex];

                // Add "Start" series (will be hidden later)
                ganttChart.NSeries.Add("B2:B6", true);
                ganttChart.NSeries[0].Name = "Start";

                // Add "Duration" series (visible bars)
                ganttChart.NSeries.Add("C2:C6", true);
                ganttChart.NSeries[1].Name = "Duration";

                // Configure chart appearance to mimic a Gantt chart
                ganttChart.GapWidth = 0; // No gap between bars

                // Hide the "Start" series by making it transparent
                ganttChart.NSeries[0].Area.ForegroundColor = Color.Transparent;
                ganttChart.NSeries[0].Border.Color = Color.Transparent;

                ganttChart.Title.Text = "Project Schedule (Gantt Chart)";

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Gantt chart created successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while generating the Gantt chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
