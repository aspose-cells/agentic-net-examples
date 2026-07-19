// Title: Detect Primary & Secondary Value Axes in Aspose.Cells Charts with Chart.HasAxis (C#)
// Description: A C# sample that builds a workbook, inserts a column chart, and uses Chart.HasAxis together with AxisType.Value and the primary/secondary flag to determine whether a value axis exists, prints the findings, and saves the workbook.
// Keywords: Aspose.Cells | Chart.HasAxis | AxisType.Value | value axis detection | primary axis check | secondary axis check | C# chart example | Aspose.Cells .NET | chart axis existence | column chart
// Common Searches: Aspose.Cells how to check for a primary value axis | Chart.HasAxis secondary axis C# example | detect value axis in Aspose.Cells chart | determine if chart has axis Aspose.Cells .NET | use AxisType.Value with HasAxis
// Developer Intent: Identify whether a chart includes a primary or secondary value axis before applying further customizations.
// Use Cases: Validate axis presence prior to setting titles, scaling, or formatting. | Add a secondary value axis only when it is missing to avoid duplication errors. | Log axis availability for dynamic report generation or debugging.
// AI Prompts: Write C# code that creates a line chart and uses Chart.HasAxis to verify both primary and secondary category axes. | Show an Aspose.Cells snippet that checks for a primary value axis and assigns a title if the axis is present. | Explain the steps for using Chart.HasAxis with AxisType.Category to detect a secondary category axis in a chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHasAxisDemo
{
    // A C# sample that builds a workbook, inserts a column chart, and uses Chart.HasAxis together with AxisType.Value and the primary/secondary flag to determine whether a value axis exists, prints the findings, and saves the workbook.
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

            // Use HasAxis to determine if a primary value axis exists
            bool hasPrimaryValueAxis = chart.HasAxis(AxisType.Value, true);
            // Use HasAxis to determine if a secondary value axis exists
            bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);

            // Output the results
            Console.WriteLine("Primary Value Axis exists: " + hasPrimaryValueAxis);
            Console.WriteLine("Secondary Value Axis exists: " + hasSecondaryValueAxis);

            // Save the workbook
            workbook.Save("ChartHasAxisDemo.xlsx");
        }
    }
}
