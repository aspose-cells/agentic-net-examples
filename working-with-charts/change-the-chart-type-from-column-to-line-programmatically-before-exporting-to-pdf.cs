// Title: Programmatically change a column chart to a line chart and export it as a PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that converts an existing column chart to a line chart in an Aspose.Cells workbook and saves the chart as a PDF file. | Write a C# snippet using Aspose.Cells to set Chart.Type from Column to Line and then call Chart.ToPdf to export only the chart. | Create a C# example that builds a column chart, switches its type to line, and uses the Aspose.Cells API to produce a PDF of the chart.
// Common Searches: Aspose.Cells C# change chart type from column to line before PDF export | How to export only a chart to PDF using Aspose.Cells .NET | C# Aspose.Cells convert column chart to line chart and save as PDF | Chart.ToPdf method example for line chart in Aspose.Cells | Modify chart type programmatically Aspose.Cells .NET
// Tags: Aspose.Cells change chart type | Aspose.Cells export chart to PDF | C# chart type conversion line | Aspose.Cells Chart.ToPdf usage | programmatic chart type modification Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartConversion
{
    // The example creates a workbook, adds sample data, inserts a column chart, changes its type to a line chart via the Chart.Type property, and then exports the chart directly to a PDF file using the Chart.ToPdf method.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart (initial type)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Change the chart type from Column to Line
            chart.Type = ChartType.Line;

            // Export the chart to a PDF file
            chart.ToPdf("LineChartOutput.pdf");

            Console.WriteLine("Chart type changed to Line and exported to PDF successfully.");
        }
    }
}
