using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class LoadWorkbookAndGetFirstChart
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the existing XLSX file
                string filePath = "input.xlsx";

                // Prevent FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet in the workbook
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count > 0)
                {
                    // Retrieve the first chart
                    Chart firstChart = worksheet.Charts[0];

                    // Output basic information about the chart
                    Console.WriteLine("First chart type: " + firstChart.Type);
                    Console.WriteLine("Chart is located on worksheet: " + firstChart.Worksheet.Name);
                }
                else
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}