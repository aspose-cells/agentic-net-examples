using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with some missing values
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("C");
                // Intentionally leave B4 empty to simulate incomplete data
                worksheet.Cells["A5"].PutValue("D");
                worksheet.Cells["B5"].PutValue(40);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Define the data range for the series
                string seriesRange = "B2:B5";
                chart.NSeries.Add(seriesRange, true);
                chart.NSeries.CategoryData = "A2:A5";

                // Evaluate data completeness for the series
                bool isComplete = true;

                // Convert the address string to a CellArea (start cell, end cell)
                string[] parts = seriesRange.Split(':');
                CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);

                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    object value = worksheet.Cells[row, area.StartColumn].Value;
                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                    {
                        isComplete = false;
                        break;
                    }
                }

                // Hide the series if data is incomplete
                chart.NSeries[0].IsFiltered = !isComplete;

                // Save the workbook
                string outputPath = "SeriesVisibilityDemo.xlsx";
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