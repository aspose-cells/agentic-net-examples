using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Load the workbook from the specified file path
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Check if the worksheet contains any charts
            if (worksheet.Charts.Count > 0)
            {
                // Retrieve the first chart in the collection
                Chart firstChart = worksheet.Charts[0];

                // Demonstrate accessing chart properties
                Console.WriteLine($"First chart type: {firstChart.Type}");
                Console.WriteLine($"Chart is located on worksheet: {firstChart.Worksheet.Name}");
            }
            else
            {
                Console.WriteLine("The worksheet does not contain any charts.");
            }
        }
    }
}