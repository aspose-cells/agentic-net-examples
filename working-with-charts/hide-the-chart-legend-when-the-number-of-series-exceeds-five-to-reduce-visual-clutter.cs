// Title: Hide chart legend in Aspose.Cells for .NET when the series count exceeds five
// Common Searches: Aspose.Cells C# hide chart legend if series count > 5 | conditionally remove legend from Excel chart using Aspose.Cells .NET | check number of series in Aspose.Cells chart and hide legend programmatically | Aspose.Cells column chart hide legend when more than five data series | C# set Legend.IsDeleted based on series count in Aspose.Cells
// Tags: chart legend conditional visibility Aspose.Cells | Aspose.Cells hide legend based on series count | Legend.IsDeleted property usage Aspose.Cells | column chart series limit handling Aspose.Cells | Excel chart generation with dynamic legend Aspose.Cells | C# Aspose.Cells chart customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartLegendConditionalHide
{
    // The example creates a workbook with six data series, adds a column chart, checks if the series count exceeds five, and hides the chart legend by setting chart.Legend.IsDeleted to true before saving the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Columns B to G – Six data series (more than five)
            for (int col = 1; col <= 6; col++) // B=1, C=2, ... G=6
            {
                // Header for each series
                sheet.Cells[0, col].PutValue($"Series {col}");
                // Random data for demonstration
                for (int row = 1; row <= 4; row++)
                {
                    sheet.Cells[row, col].PutValue(10 * col + row * 5);
                }
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add each series to the chart
            for (int col = 1; col <= 6; col++)
            {
                // Data range for the current series (e.g., B2:B5, C2:C5, ...)
                string dataRange = CellsHelper.CellIndexToName(1, col) + "2:" + CellsHelper.CellIndexToName(4, col);
                chart.NSeries.Add(dataRange, true);
            }

            // Set the category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Check the number of series and hide the legend if more than five
            if (chart.NSeries.Count > 5)
            {
                // Deleting the legend effectively hides it
                chart.Legend.IsDeleted = true;
                Console.WriteLine("Legend hidden because the chart contains more than five series.");
            }
            else
            {
                // Ensure the legend is visible when series count is five or fewer
                chart.Legend.IsDeleted = false;
                Console.WriteLine("Legend displayed (five or fewer series).");
            }

            // Save the workbook
            string outputPath = "ChartWithConditionalLegend.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
