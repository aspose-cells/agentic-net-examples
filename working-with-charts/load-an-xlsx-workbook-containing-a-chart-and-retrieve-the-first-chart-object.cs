using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Path to the existing XLSX file that contains a chart
        string filePath = "input.xlsx";

        // Load the workbook from the specified file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Check if the worksheet contains any charts
        if (worksheet.Charts.Count > 0)
        {
            // Retrieve the first chart object (uses ChartCollection indexer)
            Chart firstChart = worksheet.Charts[0];

            // Output some basic information about the chart
            Console.WriteLine("First chart retrieved successfully.");
            Console.WriteLine($"Chart Name: {firstChart.Name}");
            Console.WriteLine($"Chart Type: {firstChart.Type}");
        }
        else
        {
            Console.WriteLine("No charts found in the first worksheet.");
        }
    }
}