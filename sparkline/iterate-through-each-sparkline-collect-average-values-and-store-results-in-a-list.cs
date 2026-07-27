using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

class SparklineAverageExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (5 rows x 4 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1)); // simple numeric data
                }
            }

            // Define where the sparklines will be placed (column F)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 5,
                EndColumn = 5
            };

            // Add a sparkline group (Line type) covering the data range A1:D5
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D5", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline for each row (A1:D1, A2:D2, ..., A5:D5)
            for (int r = 0; r < 5; r++)
            {
                string rowRange = $"A{r + 1}:D{r + 1}";
                group.Sparklines.Add(rowRange, r, 5); // row index, column index (F column)
            }

            // List to hold average values of each sparkline
            List<double> sparklineAverages = new List<double>();

            // Iterate through each sparkline, compute the average of its data range, and store it
            foreach (SparklineGroup sg in sheet.SparklineGroups)
            {
                foreach (Sparkline sp in sg.Sparklines)
                {
                    // Obtain the range object from the sparkline's DataRange string
                    AsposeRange dataRange = sheet.Cells.CreateRange(sp.DataRange);

                    double sum = 0;
                    int count = 0;

                    // Accumulate numeric values in the range
                    foreach (Cell cell in dataRange)
                    {
                        if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
                        {
                            sum += val;
                            count++;
                        }
                    }

                    double average = count > 0 ? sum / count : 0;
                    sparklineAverages.Add(average);
                }
            }

            // Output the averages to console for verification
            for (int i = 0; i < sparklineAverages.Count; i++)
            {
                Console.WriteLine($"Sparkline {i + 1} average: {sparklineAverages[i]}");
            }

            // Save the workbook (ensure the directory exists)
            string outputPath = "SparklineAverageExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}