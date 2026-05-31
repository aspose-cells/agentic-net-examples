using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartValidation
{
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

                // Populate sample data with an intentional empty cell (B3)
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("B");
                // B3 left empty
                cells["A4"].PutValue("C");
                cells["B4"].PutValue(30);

                // Define the data range for the chart
                string dataRange = "B2:B4";          // Values
                string categoryRange = "A2:A4";      // Categories

                // Create a Range object covering the values area
                AsposeRange valueRange = cells.CreateRange(dataRange);

                // Validate that the range does not contain any empty cells
                bool hasEmpty = false;
                for (int row = valueRange.FirstRow; row < valueRange.FirstRow + valueRange.RowCount; row++)
                {
                    for (int col = valueRange.FirstColumn; col < valueRange.FirstColumn + valueRange.ColumnCount; col++)
                    {
                        // A cell is considered empty if its Value is null
                        if (cells[row, col].Value == null)
                        {
                            hasEmpty = true;
                            break;
                        }
                    }
                    if (hasEmpty) break;
                }

                if (hasEmpty)
                {
                    Console.WriteLine("The chart data range contains empty cells. Chart will not be created.");
                }
                else
                {
                    // Add a column chart to the worksheet
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set the data series and category data
                    chart.NSeries.Add(dataRange, true);
                    chart.NSeries.CategoryData = categoryRange;

                    // Define how to handle empty cells (not needed here as we validated)
                    chart.PlotEmptyCellsType = PlotEmptyCellsType.NotPlotted;

                    Console.WriteLine("Chart created successfully.");
                }

                // Save the workbook
                string outputPath = "ChartValidationResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}