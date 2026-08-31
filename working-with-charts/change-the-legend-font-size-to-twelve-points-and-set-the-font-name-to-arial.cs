// Title: How to set an Aspose.Cells chart legend font to Arial 12 pt using C#
// AI Prompts: Generate C# code that builds a workbook, adds a column chart with Aspose.Cells, and changes the legend font to Arial at 12 points. | Show the steps to retrieve the Legend.Font object of an Aspose.Cells chart and assign a custom font name and size in a .NET application. | Provide a complete Aspose.Cells example that creates sample data, inserts a chart, and customizes the legend typography before saving the file.
// Common Searches: Aspose.Cells C# change chart legend font to Arial 12pt | set legend font size programmatically in Aspose.Cells chart | customize legend typography for a column chart using Aspose.Cells .NET
// Tags: Aspose.Cells chart legend font styling | set legend font name Arial Aspose.Cells | legend font size 12 Aspose.Cells chart | column chart legend customization .NET | Aspose.Cells workbook chart formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendFontDemo
{
    // The example creates a workbook, populates sample data, adds a column chart, and modifies the chart legend's font to Arial with a size of 12 points, then saves the workbook as ChartWithCustomLegendFont.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the chart legend and modify its font
            // Set font name to Arial and size to 12 points
            chart.Legend.Font.Name = "Arial";
            chart.Legend.Font.Size = 12;

            // Save the workbook to a file
            workbook.Save("ChartWithCustomLegendFont.xlsx");
        }
    }
}
