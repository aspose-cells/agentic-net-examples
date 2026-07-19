// Title: Hide Chart Legend in Aspose.Cells for .NET When Series Count Exceeds Five
// Description: Shows how to create a workbook, add six data series to a column chart, and automatically hide the legend by checking the NSeries.Count and setting ShowLegend to false when more than five series are present, then save the workbook.
// Keywords: Aspose.Cells | C# chart legend | hide legend Aspose.Cells | conditional legend visibility | series count chart | ShowLegend property | Excel chart automation | .NET chart example | column chart legend | dynamic legend control
// Common Searches: Aspose.Cells hide legend if series > 5 | C# set ShowLegend based on series count | conditional chart legend Aspose.Cells .NET | how to turn off legend for many series in Excel using Aspose | auto hide chart legend Aspose.Cells
// Developer Intent: Programmatically suppress the chart legend when the number of data series in an Aspose.Cells chart exceeds a defined threshold.
// Use Cases: Generate reports with column charts that automatically hide the legend when more than five series are plotted to keep the layout clean. | Create dashboards where legend visibility toggles based on dynamic data sets, preventing overcrowded legends. | Apply the same conditional logic to other chart types (pie, line, bar) by checking NSeries.Count before rendering. | Integrate legend control into automated Excel generation pipelines for enterprise reporting.
// AI Prompts: Write C# code using Aspose.Cells that hides a chart legend when the series count is greater than five. | Provide an example that adds an arbitrary number of series to a chart and disables the legend if NSeries.Count > 5. | Explain how to adapt the conditional legend logic for pie, line, and bar charts in Aspose.Cells. | Generate a reusable method that accepts a chart object and a maximum series threshold, then sets ShowLegend accordingly.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendControl
{
    // Shows how to create a workbook, add six data series to a column chart, and automatically hide the legend by checking the NSeries.Count and setting ShowLegend to false when more than five series are present, then save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with six series (columns B to G) and four categories (rows 2 to 5)
            // Category labels
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series data
            for (int col = 1; col <= 6; col++) // Columns B (1) to G (6)
            {
                for (int row = 2; row <= 5; row++) // Rows 2 to 5
                {
                    sheet.Cells[row - 1, col].PutValue((col * 10) + row); // Arbitrary values
                }
                // Header for each series
                sheet.Cells[0, col].PutValue($"Series{col}");
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add each series to the chart (range B2:B5, C2:C5, ..., G2:G5)
            for (int col = 1; col <= 6; col++)
            {
                // Column letters for range building
                string columnLetter = CellsHelper.ColumnIndexToName(col);
                string range = $"{columnLetter}2:{columnLetter}5";
                chart.NSeries.Add(range, true);
            }

            // Set category (X) axis data (A2:A5)
            chart.NSeries.CategoryData = "A2:A5";

            // Hide legend if the number of series exceeds five
            if (chart.NSeries.Count > 5)
            {
                chart.ShowLegend = false; // Legend hidden
            }
            else
            {
                chart.ShowLegend = true; // Legend visible (default)
            }

            // Save the workbook
            workbook.Save("ChartWithConditionalLegend.xlsx");
        }
    }
}
