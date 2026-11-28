// Title: Create a column chart with a variable number of series determined by worksheet rows using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add chart series for each data row dynamically | How to generate a column chart with unknown number of series from Excel data using Aspose.Cells | Set series names from the first column when building charts with Aspose.Cells .NET | Use MaxDataRow to determine series count for a chart in Aspose.Cells C# | Programmatically build range addresses for chart series in Aspose.Cells
// Tags: dynamic chart series Aspose.Cells C# | column chart from worksheet rows | set series name from cell Aspose.Cells | use MaxDataRow to count data rows | programmatic range address Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DynamicSeriesChart
{
    // The example creates a workbook, fills category headers and series data, calculates the number of data rows with MaxDataRow, adds a column chart, sets the category range, loops through each data row to add a series and assign its name from column A, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------
                // 1. Prepare sample data
                // -----------------------------
                // Define categories (header row) in columns B to E
                string[] categories = { "Q1", "Q2", "Q3", "Q4" };
                for (int col = 0; col < categories.Length; col++)
                {
                    sheet.Cells[0, col + 1].PutValue(categories[col]); // Row 0, columns B..E
                }

                // Define series names and their values (rows 1..N)
                string[] seriesNames = { "Product A", "Product B", "Product C" };
                int[,] values = {
                    { 120, 150, 130, 170 },
                    { 80,  110, 95,  140 },
                    { 200, 180, 210, 190 }
                };

                for (int row = 0; row < seriesNames.Length; row++)
                {
                    // Series name in column A
                    sheet.Cells[row + 1, 0].PutValue(seriesNames[row]);

                    // Series values across the category columns
                    for (int col = 0; col < categories.Length; col++)
                    {
                        sheet.Cells[row + 1, col + 1].PutValue(values[row, col]);
                    }
                }

                // -----------------------------
                // 2. Determine dynamic series count
                // -----------------------------
                // MaxDataRow returns the zero‑based index of the last row that contains data.
                int lastDataRow = sheet.Cells.MaxDataRow; // includes header row (0) and series rows
                int firstSeriesRow = 1; // data starts after the header
                int seriesCount = lastDataRow - firstSeriesRow + 1;

                if (seriesCount <= 0)
                {
                    Console.WriteLine("No data rows found to create series.");
                    return;
                }

                // -----------------------------
                // 3. Add a chart to the worksheet
                // -----------------------------
                // Chart placed from row 5, column 0 to row 25, column 10
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the common category (X‑axis) data range (header row)
                // Example: "B1:E1"
                string categoryRange = $"B1:{GetColumnLetter(categories.Length + 1)}1";
                chart.NSeries.CategoryData = categoryRange;

                // -----------------------------
                // 4. Add series dynamically
                // -----------------------------
                for (int row = firstSeriesRow; row <= lastDataRow; row++)
                {
                    // Build the address for the current row's values, e.g., "B2:E2"
                    string seriesRange = $"B{row + 1}:{GetColumnLetter(categories.Length + 1)}{row + 1}";

                    // Add the series; false indicates that the data is plotted by row
                    chart.NSeries.Add(seriesRange, false);

                    // Retrieve the just‑added series (it will be the last one in the collection)
                    Series addedSeries = chart.NSeries[chart.NSeries.Count - 1];

                    // Set the series name from column A of the same row
                    addedSeries.Name = sheet.Cells[row, 0].StringValue;
                }

                // Optional: give the chart a title
                chart.Title.Text = "Dynamic Series Chart";

                // -----------------------------
                // 5. Save the workbook
                // -----------------------------
                string outputPath = "DynamicSeriesChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static string GetColumnLetter(int columnNumber)
        {
            if (columnNumber < 1)
                throw new ArgumentOutOfRangeException(nameof(columnNumber));

            string columnLetter = string.Empty;
            while (columnNumber > 0)
            {
                int remainder = (columnNumber - 1) % 26;
                columnLetter = (char)(65 + remainder) + columnLetter;
                columnNumber = (columnNumber - remainder) / 26;
            }
            return columnLetter;
        }
    }
}
