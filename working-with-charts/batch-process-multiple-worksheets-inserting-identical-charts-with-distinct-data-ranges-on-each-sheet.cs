// Title: Create identical clustered column charts on each worksheet with individual data ranges using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add same chart to every worksheet with different data range | how to generate batch charts across multiple sheets in a .xlsx using Aspose.Cells | set per-sheet chart source data programmatically with Aspose.Cells .NET | loop through worksheets to insert column chart in Aspose.Cells C# example | save workbook after adding charts to each sheet Aspose.Cells
// Tags: add chart to each worksheet Aspose.Cells | per-sheet chart data range Aspose.Cells | clustered column chart generation .NET | batch chart insertion workbook Aspose.Cells | populate worksheet data before chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchChartExample
{
    // The example creates a new workbook, adds three worksheets, fills cells A1:B5 on each sheet with distinct values, inserts a clustered column chart on each sheet that references its own A1:B5 range, sets a sheet‑specific title and legend position, and saves the workbook as BatchChartsOutput.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook with three worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear(); // remove the default sheet
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Iterate over each worksheet to insert identical charts with sheet‑specific data
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    Cells cells = sheet.Cells;

                    // Populate sample data in A1:B5 (different values per sheet)
                    for (int row = 0; row < 5; row++)
                    {
                        // Category (e.g., Q1, Q2, ...) in column A
                        cells[row, 0].PutValue($"Q{row + 1}");
                        // Numeric value in column B – vary by sheet index to make data distinct
                        cells[row, 1].PutValue((row + 1) * (i + 1) * 10);
                    }

                    // Add a clustered column chart to the current sheet
                    // Position: from row 7, column 0 to row 25, column 10
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set the data range for the series (A1:B5) – each sheet uses its own data
                    chart.NSeries.Add("A1:B5", true);

                    // Optional: set chart title to identify the sheet
                    chart.Title.Text = $"Sales Data - {sheet.Name}";
                    chart.Title.Font.IsBold = true;

                    // Optional: set legend position
                    chart.Legend.Position = LegendPositionType.Right;
                }

                // Define output file path
                string outputPath = "BatchChartsOutput.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
