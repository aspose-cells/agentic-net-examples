// Title: C# – Load an XLSX workbook and retrieve the first chart using Aspose.Cells
// Description: Loads an XLSX file from a given path, accesses the first worksheet, checks its ChartCollection, and returns the first chart (index 0) with its name, type, and parent worksheet. Includes file‑existence verification and comprehensive exception handling.
// Keywords: Aspose.Cells load workbook C# | read XLSX chart Aspose.Cells | first chart from worksheet | chart collection Aspose.Cells | .NET chart metadata extraction | file not found handling Aspose.Cells
// Common Searches: how to load an xlsx file and get the first chart in C# | aspnet retrieve chart name and type from worksheet | aspose.cells check if worksheet contains charts | c# example loading workbook and accessing charts
// Developer Intent: Load a workbook from a file path and obtain the first chart object on the first worksheet.
// Use Cases: Quickly display basic chart information (name, type, sheet) for validation in reporting tools. | Confirm that a workbook contains at least one chart before performing batch chart processing. | Extract chart metadata for documentation, audit logs, or downstream data pipelines.
// AI Prompts: Generate a C# method that loads an XLSX file, returns the first chart object, and handles missing files and empty chart collections gracefully. | Write code to iterate over all charts in a workbook and output each chart's name, type, and worksheet using Aspose.Cells. | Create robust exception handling for workbook loading and chart access, providing clear messages for file‑not‑found and no‑chart scenarios.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Loads an XLSX file from a given path, accesses the first worksheet, checks its ChartCollection, and returns the first chart (index 0) with its name, type, and parent worksheet. Includes file‑existence verification and comprehensive exception handling.
    public class LoadWorkbookAndGetFirstChart
    {
        public static void Run(string filePath)
        {
            try
            {
                // Load the workbook from the specified XLSX file path
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet in the workbook
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the chart collection from the worksheet
                ChartCollection charts = worksheet.Charts;

                // Check if there is at least one chart in the collection
                if (charts.Count > 0)
                {
                    // Obtain the first chart (index 0)
                    Chart firstChart = charts[0];

                    // Display some properties of the chart
                    Console.WriteLine("First chart name: " + firstChart.Name);
                    Console.WriteLine("First chart type: " + firstChart.Type);
                    Console.WriteLine("Chart belongs to worksheet: " + firstChart.Worksheet.Name);
                }
                else
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Determine the file path: use argument if provided, otherwise default to "sample.xlsx"
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Execute the example
            LoadWorkbookAndGetFirstChart.Run(filePath);
        }
    }
}
