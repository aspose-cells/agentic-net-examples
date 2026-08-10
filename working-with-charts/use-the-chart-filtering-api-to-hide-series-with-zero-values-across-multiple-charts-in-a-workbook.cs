// Title: Hide Zero‑Value Series in Multiple Aspose.Cells Charts (C#) Using IsFiltered
// Description: This example creates a workbook, populates it with categories and three data series (one containing only zeros), adds a column chart and a line chart, detects the all‑zero series with a helper method, sets the series' IsFiltered property to hide it, and saves the file as FilteredCharts.xlsx.
// Keywords: Aspose.Cells chart filtering | C# hide zero series | IsFiltered property | multiple Excel charts | remove empty series Aspose.Cells | .NET chart API | Excel workbook automation | Aspose.Cells example
// Common Searches: Aspose.Cells hide series with zero values | set IsFiltered true for chart series C# | filter out empty series in Excel chart using Aspose | chart series filtering API Aspose.Cells | remove zero data series from multiple charts
// Developer Intent: Programmatically hide any chart series that consist solely of zero or blank values across all charts in an Aspose.Cells workbook.
// Use Cases: Financial dashboards where categories with no activity should not appear in column or line charts. | Automated report generation that cleans up charts by omitting empty data series before distribution. | Template‑based workbook processing that applies a uniform zero‑value filter to every chart.
// AI Prompts: Generate C# code that iterates through all charts in an Aspose.Cells workbook and hides series whose data range contains only zeros. | Provide a reusable method to check if a cell range is all zero and apply the IsFiltered flag to the corresponding series for each chart. | Explain the impact of the IsFiltered property on chart rendering in Excel files created with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartFilteringDemo
{
    // This example creates a workbook, populates it with categories and three data series (one containing only zeros), adds a column chart and a line chart, detects the all‑zero series with a helper method, sets the series' IsFiltered property to hide it, and saves the file as FilteredCharts.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Sample data for two charts (each chart has two series)
                // -------------------------------------------------
                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");

                // Data rows
                string[] categories = { "A", "B", "C", "D" };
                double[,] values = {
                    { 10, 0, 5 },   // Row 2
                    { 20, 0, 15 },  // Row 3
                    { 30, 0, 25 },  // Row 4
                    { 40, 0, 35 }   // Row 5
                };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[$"A{i + 2}"].PutValue(categories[i]);
                    sheet.Cells[$"B{i + 2}"].PutValue(values[i, 0]); // Series1 (non‑zero)
                    sheet.Cells[$"C{i + 2}"].PutValue(values[i, 1]); // Series2 (all zeros)
                    sheet.Cells[$"D{i + 2}"].PutValue(values[i, 2]); // Series3 (non‑zero)
                }

                // -------------------------------------------------
                // Add first chart (Column)
                // -------------------------------------------------
                int chartIdx1 = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 15);
                Chart chart1 = sheet.Charts[chartIdx1];
                chart1.Title.Text = "First Chart";

                // Add series
                int s1c1Idx = chart1.NSeries.Add("B2:B5", true);
                Series s1c1 = chart1.NSeries[s1c1Idx];
                int s2c1Idx = chart1.NSeries.Add("C2:C5", true);
                Series s2c1 = chart1.NSeries[s2c1Idx];
                int s3c1Idx = chart1.NSeries.Add("D2:D5", true);
                Series s3c1 = chart1.NSeries[s3c1Idx];
                chart1.NSeries.CategoryData = "A2:A5";

                // Hide series that are all zero
                if (IsSeriesAllZero(sheet, "C2:C5"))
                    s2c1.IsFiltered = true;

                // -------------------------------------------------
                // Add second chart (Line)
                // -------------------------------------------------
                int chartIdx2 = sheet.Charts.Add(ChartType.Line, 25, 0, 40, 15);
                Chart chart2 = sheet.Charts[chartIdx2];
                chart2.Title.Text = "Second Chart";

                int s1c2Idx = chart2.NSeries.Add("B2:B5", true);
                Series s1c2 = chart2.NSeries[s1c2Idx];
                int s2c2Idx = chart2.NSeries.Add("C2:C5", true);
                Series s2c2 = chart2.NSeries[s2c2Idx];
                chart2.NSeries.CategoryData = "A2:A5";

                if (IsSeriesAllZero(sheet, "C2:C5"))
                    s2c2.IsFiltered = true;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "FilteredCharts.xlsx";

                // Ensure the directory exists (prevents FileNotFoundException)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Checks whether all numeric values in the specified range are zero (or empty)
        private static bool IsSeriesAllZero(Worksheet ws, string rangeRef)
        {
            // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range range = ws.Cells.CreateRange(rangeRef);
            int startRow = range.FirstRow;
            int endRow = startRow + range.RowCount - 1;
            int startCol = range.FirstColumn;
            int endCol = startCol + range.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = ws.Cells[row, col];
                    double value = 0;

                    if (cell.Type == CellValueType.IsNumeric)
                        value = cell.DoubleValue;
                    else if (cell.Type == CellValueType.IsString && double.TryParse(cell.StringValue, out double parsed))
                        value = parsed;

                    if (value != 0)
                        return false; // Found a non‑zero value.
                }
            }
            return true; // All cells are zero or empty.
        }
    }
}
