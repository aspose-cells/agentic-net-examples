// Title: How to use Aspose.Cells Chart.HasAxis to verify primary and secondary value axes in a C# column chart
// Common Searches: Aspose.Cells C# check if a column chart contains a primary value axis | Determine secondary value axis existence with Chart.HasAxis in .NET | C# example of chart.HasAxis method for column charts using Aspose.Cells | Verify whether an Excel chart has value axes using Aspose.Cells API | Use AxisType.Value with chart.HasAxis to detect chart axes in C#
// Tags: Aspose.Cells chart.HasAxis value axis detection | C# column chart axis existence verification | primary secondary value axis check Aspose | Excel chart axis presence detection .NET | chart data range setup and axis check Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartHasAxisDemo
{
    // The example creates a workbook, adds sample category and value data, inserts a column chart, and uses Chart.HasAxis with AxisType.Value for both primary and secondary axes to determine their presence. It prints the results to the console and saves the workbook as ChartHasAxisDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart (including both categories and values)
                chart.SetChartDataRange("A1:B4", true);

                // Use HasAxis to determine whether a primary and secondary value axis exist
                bool hasPrimaryValueAxis = chart.HasAxis(AxisType.Value, true);
                bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);

                // Output the results to the console
                Console.WriteLine("Primary Value Axis exists: " + hasPrimaryValueAxis);
                Console.WriteLine("Secondary Value Axis exists: " + hasSecondaryValueAxis);

                // Save the workbook to a file
                string outputPath = "ChartHasAxisDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                // Print any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
