// Title: Use Chart.HasAxis to detect primary and secondary value axes on an Aspose.Cells column chart in C#
// AI Prompts: Write a C# console program that creates a column chart with sample data and calls Chart.HasAxis to report whether the primary value axis exists. | Show how to invoke Chart.HasAxis for the secondary value axis on a chart built with Aspose.Cells and output the result to the console. | Generate C# code that sets a chart data range, checks both primary and secondary value axes using Chart.HasAxis, and saves the workbook.
// Common Searches: asp.net aspose.cells chart.hasaxis example for primary value axis | c# how to check secondary value axis in an Aspose.Cells chart | using Chart.HasAxis to verify axes on a column chart with Aspose.Cells | determine if a chart has a value axis in Aspose.Cells .NET | chart.HasAxis method usage Aspose.Cells C# tutorial
// Tags: Aspose.Cells Chart.HasAxis value axis detection | C# column chart primary axis verification | C# column chart secondary axis verification | Aspose.Cells set chart data range C# | Aspose.Cells workbook save example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHasAxisDemo
{
    // The program creates a workbook, adds a column chart with sample data, uses Chart.HasAxis to determine whether the primary and secondary value axes are present, prints the results, and saves the file as ChartHasAxisDemo.xlsx.
    class Program
    {
        static void Main()
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

            // Check if the primary value axis exists
            bool hasPrimaryValueAxis = chart.HasAxis(AxisType.Value, true);
            // Check if the secondary value axis exists
            bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);

            // Output the results
            Console.WriteLine("Primary Value Axis exists: " + hasPrimaryValueAxis);
            Console.WriteLine("Secondary Value Axis exists: " + hasSecondaryValueAxis);

            // Save the workbook
            workbook.Save("ChartHasAxisDemo.xlsx");
        }
    }
}
