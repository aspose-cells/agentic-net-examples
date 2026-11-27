// Title: Validate chart data range for empty cells before creating a chart with Aspose.Cells in C#
// Common Searches: C# code to check for blank cells in a chart data range using Aspose.Cells | how to prevent Aspose.Cells chart generation when source range contains empty cells | validate Excel range for null values before adding a column chart with Aspose.Cells .NET | Aspose.Cells chart creation abort on empty cells in data series C# example | detect empty cells in worksheet range prior to charting with Aspose.Cells
// Tags: chart data range validation Aspose.Cells | empty cell detection before chart creation C# | Aspose.Cells NSeries source range check | skip chart generation on blank data Aspose.Cells | column chart with validated range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartDataRangeValidator
{
    // The example loads or creates a workbook, scans the specified cell range for null values, aborts chart creation if any empty cells are found, and otherwise adds a column chart using the clean range, sets category data, names the series, and saves the workbook as ChartWithValidatedData.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Load or create a workbook with data that will be used for the chart.
                // -----------------------------------------------------------------
                Workbook workbook = null;
                string sourceFile = args.Length > 0 ? args[0] : null;

                if (!string.IsNullOrEmpty(sourceFile))
                {
                    // Verify that the file exists before attempting to load it.
                    if (File.Exists(sourceFile))
                    {
                        workbook = new Workbook(sourceFile);
                        Console.WriteLine($"Loaded workbook from \"{sourceFile}\".");
                    }
                    else
                    {
                        Console.WriteLine($"File \"{sourceFile}\" not found. Creating a workbook with sample data instead.");
                        workbook = CreateSampleWorkbook();
                    }
                }
                else
                {
                    workbook = CreateSampleWorkbook();
                }

                // -----------------------------------------------------------------
                // 2. Determine the chart data range to validate.
                // -----------------------------------------------------------------
                string dataRange = args.Length > 1 ? args[1] : "A2:B5"; // Default range excludes header row.
                Console.WriteLine($"Validating data range \"{dataRange}\".");

                // Obtain the first worksheet (the one containing the data).
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a Range object from the address string (use Aspose.Cells.Range to avoid ambiguity).
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(dataRange);

                // -----------------------------------------------------------------
                // 3. Validate that the range does not contain empty cells.
                // -----------------------------------------------------------------
                bool hasEmptyCells = false;
                int startRow = range.FirstRow;
                int endRow = startRow + range.RowCount - 1;
                int startCol = range.FirstColumn;
                int endCol = startCol + range.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        // Treat a cell as empty when its Value is null.
                        if (worksheet.Cells[row, col].Value == null)
                        {
                            hasEmptyCells = true;
                            Console.WriteLine($"Empty cell found at {CellIndexToName(row, col)} within the range.");
                        }
                    }
                }

                if (hasEmptyCells)
                {
                    Console.WriteLine("Chart creation aborted because the data range contains empty cells.");
                }
                else
                {
                    // -----------------------------------------------------------------
                    // 4. Create the chart because the data range is clean.
                    // -----------------------------------------------------------------
                    // Define chart position on the worksheet (rows/columns are zero‑based).
                    int chartTopRow = 7;
                    int chartLeftColumn = 0;
                    int chartBottomRow = 25;
                    int chartRightColumn = 10;

                    int chartIndex = worksheet.Charts.Add(
                        ChartType.Column, chartTopRow, chartLeftColumn, chartBottomRow, chartRightColumn);
                    Chart chart = worksheet.Charts[chartIndex];

                    // Set the data source for the chart.
                    chart.NSeries.Add(dataRange, true);

                    // Optionally set the category axis data explicitly (useful when the range includes headers).
                    string categoryRange = $"{CellIndexToName(startRow, startCol)}:{CellIndexToName(endRow, startCol)}";
                    chart.NSeries.CategoryData = categoryRange;

                    // Give the series a friendly name.
                    if (chart.NSeries.Count > 0)
                    {
                        chart.NSeries[0].Name = "Sample Series";
                    }

                    // Save the workbook with the chart.
                    string outputPath = "ChartWithValidatedData.xlsx";
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Chart created successfully and saved to \"{outputPath}\".");
                }
            }
            catch (Exception ex)
            {
                // Any unexpected error is reported to the console.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        // Creates a workbook populated with sample data.
        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Q1");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["A3"].PutValue("Q2");
            ws.Cells["B3"].PutValue(150);
            ws.Cells["A4"].PutValue("Q3");
            ws.Cells["B4"].PutValue(180);
            ws.Cells["A5"].PutValue("Q4");
            ws.Cells["B5"].PutValue(200);

            Console.WriteLine("Created a workbook with sample data.");
            return wb;
        }

        private static string CellIndexToName(int rowIndex, int columnIndex)
        {
            // Aspose.Cells provides a utility method for this conversion.
            return CellsHelper.CellIndexToName(rowIndex, columnIndex);
        }
    }
}
