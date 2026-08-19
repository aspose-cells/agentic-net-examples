// Title: C# – Detect Primary and Secondary Value Axes with Aspose.Cells Chart.HasAxis
// Description: This example creates a workbook, adds sample data, inserts a column chart, and uses the Chart.HasAxis method (AxisType.Value) to verify the presence of primary and secondary value axes. The results are printed to the console and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells Chart.HasAxis C# | check chart value axis Aspose | primary value axis detection | secondary value axis detection | Aspose.Cells axis existence | C# chart axis verification | Aspose.Cells chart API
// Common Searches: Aspose.Cells how to check if a chart has a value axis in C# | Chart.HasAxis example for primary and secondary axes | C# code to detect secondary value axis in Aspose.Cells | Determine chart axes presence with Aspose.Cells | Aspose.Cells chart axis verification tutorial
// Developer Intent: Identify whether a chart contains primary and/or secondary value axes using the Chart.HasAxis method.
// Use Cases: Validate chart layout before exporting to PDF or image formats. | Add a secondary value axis only when it does not already exist. | Log axis configuration for each chart when generating automated reports.
// AI Prompts: Write C# code that adds a secondary value axis to an Aspose.Cells chart only if Chart.HasAxis(AxisType.Value, false) returns false. | Provide an Aspose.Cells snippet that loops through all charts in a workbook and reports which axes (category, value, series) are present. | Explain the difference between primary and secondary axes in Aspose.Cells and show how to query them with HasAxis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds sample data, inserts a column chart, and uses the Chart.HasAxis method (AxisType.Value) to verify the presence of primary and secondary value axes. The results are printed to the console and the workbook is saved as an XLSX file.
    public class ChartHasAxisDemo
    {
        public static void Run()
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

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Use HasAxis to determine if a primary and secondary value axis exist
                bool hasPrimaryValueAxis = chart.HasAxis(AxisType.Value, true);
                bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);

                // Output the results
                Console.WriteLine("Primary Value Axis exists: " + hasPrimaryValueAxis);
                Console.WriteLine("Secondary Value Axis exists: " + hasSecondaryValueAxis);

                // Save the workbook
                string outputPath = "ChartHasAxisDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
