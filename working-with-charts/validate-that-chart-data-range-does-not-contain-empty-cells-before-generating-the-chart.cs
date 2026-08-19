// Title: Validate Chart Data Range for Empty Cells Before Creating a Column Chart – Aspose.Cells for .NET
// Description: C# example that creates a workbook, inserts sample data with a deliberate blank cell, checks the series range for null values using a custom IsRangeWithoutEmptyCells method (Range.IsBlank and cell‑by‑cell inspection), aborts chart creation if any cell is empty, otherwise adds a column chart, sets series and category ranges, applies PlotEmptyCellsType.NotPlotted, and saves the file as ValidatedChart.xlsx.
// Keywords: Aspose.Cells chart validation | C# empty cells range check | Range.IsBlank Aspose.Cells | prevent chart creation blank data | PlotEmptyCellsType NotPlotted | column chart Aspose.Cells .NET | worksheet range null detection | Aspose.Cells example C#
// Common Searches: Aspose.Cells check for empty cells before adding a chart | C# validate series range does not contain blanks Aspose.Cells | how to skip chart generation when data contains null values | detect blank cells in worksheet range Aspose.Cells .NET | set PlotEmptyCellsType after range validation
// Developer Intent: Confirm that every cell in the chart's data series is populated before generating the chart to avoid errors or unwanted empty points.
// Use Cases: Abort chart creation when any cell in the series range is empty, preventing runtime exceptions. | Automatically apply PlotEmptyCellsType.NotPlotted only after the data range is verified as complete. | Log or display a warning message and skip chart generation if missing values are detected.
// AI Prompts: Write a C# method using Aspose.Cells that returns true only if a given worksheet range contains no empty cells. | Generate Aspose.Cells code that creates a line chart after confirming both series and category ranges are free of null values. | Show how to handle missing data in a chart by validating the range first and then setting PlotEmptyCellsType to NotPlotted.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartValidation
{
    // C# example that creates a workbook, inserts sample data with a deliberate blank cell, checks the series range for null values using a custom IsRangeWithoutEmptyCells method (Range.IsBlank and cell‑by‑cell inspection), aborts chart creation if any cell is empty, otherwise adds a column chart, sets series and category ranges, applies PlotEmptyCellsType.NotPlotted, and saves the file as ValidatedChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with an empty cell in the series range
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                // B3 is intentionally left empty
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Define the data range for the chart series
                string seriesRange = "B2:B4";

                // Validate that the defined range does not contain empty cells
                if (IsRangeWithoutEmptyCells(sheet, seriesRange))
                {
                    // Add a chart (lifecycle rule: creation and later saving)
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set the series data range
                    chart.NSeries.Add(seriesRange, true);
                    chart.NSeries.CategoryData = "A2:A4";

                    // Optional: define how to handle empty cells (not needed after validation)
                    chart.PlotEmptyCellsType = PlotEmptyCellsType.NotPlotted;
                }
                else
                {
                    Console.WriteLine($"The range \"{seriesRange}\" contains empty cells. Chart creation aborted.");
                }

                // Save the workbook (save rule)
                workbook.Save("ValidatedChart.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as ValidatedChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        static bool IsRangeWithoutEmptyCells(Worksheet sheet, string rangeAddress)
        {
            // Create a Range object for the address using the Aspose.Cells alias to avoid ambiguity
            AsposeRange range = sheet.Cells.CreateRange(rangeAddress);

            // If the whole range is blank, it definitely contains empty cells
            if (range.IsBlank())
                return false;

            // Iterate through each cell in the range to detect blanks
            for (int row = 0; row < range.RowCount; row++)
            {
                for (int col = 0; col < range.ColumnCount; col++)
                {
                    int actualRow = range.FirstRow + row;
                    int actualCol = range.FirstColumn + col;

                    // A cell is considered empty if its value is null
                    if (sheet.Cells[actualRow, actualCol].Value == null)
                        return false;
                }
            }

            // No empty cells found
            return true;
        }
    }
}
