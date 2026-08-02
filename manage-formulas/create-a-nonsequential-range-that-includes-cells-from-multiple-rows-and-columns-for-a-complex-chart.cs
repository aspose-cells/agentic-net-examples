// Title: Create a Union (non‑sequential) Range for a Column Chart with Aspose.Cells .NET
// Description: Demonstrates how to build a workbook, fill column A with categories and column C with values, combine the two ranges into a UnionRange, assign it to a column chart using SetChartDataRange, add a title, and save the file as XLSX.
// Keywords: Aspose.Cells UnionRange | non‑contiguous chart data .NET | SetChartDataRange example | column chart from multiple columns | Aspose.Cells chart data source
// Common Searches: Aspose.Cells combine A1:A5 and C1:C5 for chart | create union range for chart Aspose .NET | SetChartDataRange non‑contiguous cells | column chart using multiple columns Aspose.Cells
// Developer Intent: Generate a non‑sequential range that merges cells from different columns and use it as the data source for a column chart.
// Use Cases: Display categories in column A and values in column C on a single column chart. | Produce financial or sales reports that pull data from separate columns into one visual chart. | Reuse the same UnionRange to feed several chart types (e.g., column, line, pie) within the same worksheet.
// AI Prompts: Show C# code to create a UnionRange from A1:A5 and C1:C5 and bind it to a column chart with Aspose.Cells. | Provide an example of adding a line chart that uses a non‑contiguous data range in Aspose.Cells for .NET. | Explain how to control series plotting direction when using SetChartDataRange with a UnionRange.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsComplexChartDemo
{
    // Demonstrates how to build a workbook, fill column A with categories and column C with values, combine the two ranges into a UnionRange, assign it to a column chart using SetChartDataRange, add a title, and save the file as XLSX.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data in two non‑contiguous columns (A and C)
                // Column A will be categories, Column C will be values
                for (int i = 0; i < 5; i++)
                {
                    // Category labels in A1:A5
                    dataSheet.Cells[i, 0].PutValue($"Cat {i + 1}");
                    // Corresponding values in C1:C5
                    dataSheet.Cells[i, 2].PutValue((i + 1) * 10);
                }

                // Create two separate ranges: A1:A5 and C1:C5
                AsposeRange rangeA = dataSheet.Cells.CreateRange("A1", "A5");
                AsposeRange rangeC = dataSheet.Cells.CreateRange("C1", "C5");

                // Combine the two ranges into a non‑sequential (union) range
                UnionRange unionRange = rangeA.UnionRanges(new AsposeRange[] { rangeC });

                // Add a column chart to the same worksheet
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = dataSheet.Charts[chartIndex];

                // Set the chart's data source to the union range.
                // The second parameter (true) indicates that series are plotted by column.
                chart.SetChartDataRange(unionRange.RefersTo, true);

                // Optionally set chart title
                chart.Title.Text = "Complex Chart with Non‑Sequential Range";

                // Save the workbook
                workbook.Save("ComplexChartWithUnionRange.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
