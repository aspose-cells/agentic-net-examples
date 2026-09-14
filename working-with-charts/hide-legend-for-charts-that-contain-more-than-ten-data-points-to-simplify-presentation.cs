// Title: Automatically hide the legend of a column chart in Aspose.Cells for .NET when the series exceeds ten data points (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a column chart, determines the series length, and disables the legend when the length exceeds ten. | Demonstrate how to use the CellArea class to evaluate a chart's data range and conditionally set ShowLegend to false in an Excel workbook.
// Common Searches: Aspose.Cells C# hide chart legend for series with more than 10 entries | conditional legend visibility in Excel column chart using Aspose.Cells .NET | count data points in Aspose.Cells chart series to control legend display | programmatically remove legend from large column chart in Aspose.Cells | how to suppress Excel chart legend when data range exceeds ten rows Aspose.Cells
// Tags: Aspose.Cells legend visibility control | C# chart series length detection | Conditional column chart legend | CellArea range parsing Aspose | Excel chart legend suppression

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendControl
{
    // The example creates a workbook, fills cells A2:A13 and B2:B13 with sample data, adds a column chart, uses CellArea to calculate the number of data points in the series, hides the chart legend when the count is greater than ten, and saves the file as ChartWithConditionalLegend.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with more than ten data points (B2:B13)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 13; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 5); // arbitrary numeric values
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B13", true);
            chart.NSeries.CategoryData = "A2:A13";

            // Determine the number of data points in the series range
            // Using CellArea to parse the range string
            CellArea dataArea = CellArea.CreateCellArea("B2", "B13");
            int dataPointCount = dataArea.EndRow - dataArea.StartRow + 1; // inclusive count

            // Hide the legend if there are more than ten data points
            if (dataPointCount > 10)
            {
                chart.ShowLegend = false;
            }

            // Save the workbook
            workbook.Save("ChartWithConditionalLegend.xlsx", SaveFormat.Xlsx);
        }
    }
}
