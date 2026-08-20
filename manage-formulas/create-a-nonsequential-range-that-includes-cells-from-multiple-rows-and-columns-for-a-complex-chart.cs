// Title: Aspose.Cells .NET: Build a Non‑Contiguous Range and Apply It to a Column Chart
// Description: Demonstrates how to create two separate ranges (A1:B5 and D1:E5), merge them with the UnionRanges method into a non‑sequential range, and use the combined address (e.g., "A1:B5,D1:E5") as the data source for a column chart in a workbook saved as an Excel file.
// Keywords: Aspose.Cells UnionRanges | non‑contiguous range .NET | Aspose.Cells chart data source | column chart from multiple ranges | C# Aspose.Cells example | Excel non‑sequential range | Aspose.Cells chart series address
// Common Searches: Aspose.Cells create non‑contiguous range for chart | UnionRanges method C# example | how to bind multiple ranges to a chart in Aspose.Cells | column chart with A1:B5 and D1:E5 data | Aspose.Cells combine separate data blocks
// Developer Intent: Merge distinct cell blocks into a single non‑contiguous range and use that range as the series source for a column chart.
// Use Cases: Combine sales data stored in separate tables (e.g., columns A‑B and D‑E) into one visual chart without moving cells. | Create a dashboard where category labels and values are placed in different sections but need a unified chart representation. | Generate Excel reports that pull data from scattered ranges and display them together in a single chart.
// AI Prompts: Show how to use Aspose.Cells UnionRanges to merge A1:B5 and D1:E5 and assign the result to a chart series in C#. | Provide a C# snippet that creates a non‑contiguous range, adds a column chart, and sets the series formula using the union range address. | Explain how to retrieve the address string of a UnionRange and use it with Chart.NSeries.Add in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsComplexChart
{
    // Demonstrates how to create two separate ranges (A1:B5 and D1:E5), merge them with the UnionRanges method into a non‑sequential range, and use the combined address (e.g., "A1:B5,D1:E5") as the data source for a column chart in a workbook saved as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate first data block (A1:B5)
                for (int row = 0; row < 5; row++)
                {
                    cells[row, 0].PutValue($"Item {row + 1}");   // Column A - categories
                    cells[row, 1].PutValue((row + 1) * 10);     // Column B - values
                }

                // Populate second data block (D1:E5) – non‑sequential with respect to the first block
                for (int row = 0; row < 5; row++)
                {
                    cells[row, 3].PutValue($"Group {row + 1}"); // Column D - categories
                    cells[row, 4].PutValue((row + 1) * 15);     // Column E - values
                }

                // Create two separate ranges
                AsposeRange range1 = cells.CreateRange("A1", "B5"); // First block
                AsposeRange range2 = cells.CreateRange("D1", "E5"); // Second block

                // Union the two ranges into a non‑sequential range
                UnionRange unionRange = range1.UnionRanges(new AsposeRange[] { range2 });

                // Add a column chart to the worksheet (positioned below the data)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Use the union range as the data source for the chart
                // unionRange.RefersTo returns the address string like "A1:B5,D1:E5"
                chart.NSeries.Add(unionRange.RefersTo, true);

                // Optional: set a title for clarity
                chart.Title.Text = "Complex Chart with Non‑Sequential Data";

                // Save the workbook
                workbook.Save("ComplexChartWithNonSequentialRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
