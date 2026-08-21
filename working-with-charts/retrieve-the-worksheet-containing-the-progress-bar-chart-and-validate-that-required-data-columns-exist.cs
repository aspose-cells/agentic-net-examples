// Title: C# – Find the Progress Bar chart worksheet and verify required columns using Aspose.Cells
// Description: Load an XLSX workbook with Aspose.Cells, search each sheet for a Bar or Column chart whose title contains “Progress Bar”, retrieve the chart's parent worksheet, and confirm that the first row includes the mandatory “Task” and “Progress” headers. Missing headers are logged and the workbook is saved.
// Keywords: Aspose.Cells C# chart search | find chart by title Aspose | retrieve worksheet from chart Aspose.Cells | validate Excel headers C# | progress bar chart validation | Excel column header check | LoadOptions Aspose.Cells | C# Excel automation | Aspose.Cells example GitHub | code snippet Aspose.Cells
// Common Searches: Aspose.Cells locate chart by title C# | How to get worksheet of a specific chart in Excel using Aspose.Cells | Check for required column headers in Excel with Aspose.Cells | Validate progress bar chart data before saving with Aspose.Cells | C# code to find Progress Bar chart and verify Task and Progress columns
// Developer Intent: Identify the worksheet that contains a Progress Bar chart and ensure it has the required Task and Progress columns.
// Use Cases: Automated quality‑check of reporting templates that must include a progress bar visual and specific data columns. | Pre‑publish validation of Excel workbooks to guarantee chart data integrity and header presence. | Batch processing of multiple files to confirm required columns before downstream analytics.
// AI Prompts: Generate C# code with Aspose.Cells that finds a chart whose title contains 'Progress Bar' and returns the worksheet name. | Write a method that scans the first row of a worksheet for 'Task' and 'Progress' headers and logs any missing columns. | Explain how LoadOptions.CheckDataValid influences workbook loading in Aspose.Cells. | Adapt the example to work with .xls files and a Pie chart type.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartValidation
{
    // Load an XLSX workbook with Aspose.Cells, search each sheet for a Bar or Column chart whose title contains “Progress Bar”, retrieve the chart's parent worksheet, and confirm that the first row includes the mandatory “Task” and “Progress” headers. Missing headers are logged and the workbook is saved.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                Validator.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }

    public class Validator
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook with data validation checking enabled
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CheckDataValid = true
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            Chart progressBarChart = null;
            Worksheet chartWorksheet = null;

            // Locate the progress bar chart
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart chart in ws.Charts)
                {
                    if (chart.Type == ChartType.Bar || chart.Type == ChartType.Column)
                    {
                        if (chart.Title != null && !string.IsNullOrEmpty(chart.Title.Text) &&
                            chart.Title.Text.Contains("Progress Bar", StringComparison.OrdinalIgnoreCase))
                        {
                            progressBarChart = chart;
                            chartWorksheet = chart.Worksheet;
                            break;
                        }
                    }
                }
                if (progressBarChart != null) break;
            }

            if (progressBarChart == null)
            {
                Console.WriteLine("Progress Bar chart not found in the workbook.");
                return;
            }

            Console.WriteLine($"Progress Bar chart is located in worksheet: {chartWorksheet.Name}");

            // Verify required column headers
            string[] requiredHeaders = { "Task", "Progress" };
            Cells cells = chartWorksheet.Cells;
            int lastColumn = cells.MaxColumn;

            foreach (string header in requiredHeaders)
            {
                bool found = false;
                for (int col = 0; col <= lastColumn; col++)
                {
                    if (string.Equals(cells[0, col].StringValue, header, StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    Console.WriteLine($"Required column \"{header}\" is missing in worksheet \"{chartWorksheet.Name}\".");
                }
                else
                {
                    Console.WriteLine($"Column \"{header}\" exists.");
                }
            }

            // Save the workbook (if any modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
    }
}
