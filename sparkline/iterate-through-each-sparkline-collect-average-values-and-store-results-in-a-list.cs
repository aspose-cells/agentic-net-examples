using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (two rows of values)
            worksheet.Cells["A1"].PutValue(5);
            worksheet.Cells["B1"].PutValue(2);
            worksheet.Cells["C1"].PutValue(1);
            worksheet.Cells["D1"].PutValue(3);
            worksheet.Cells["A2"].PutValue(7);
            worksheet.Cells["B2"].PutValue(4);
            worksheet.Cells["C2"].PutValue(6);
            worksheet.Cells["D2"].PutValue(8);

            // Define where the sparklines will be placed
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 1,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group that uses the data range A1:D2
            int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D2", false, location);
            SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

            // List to hold the average of each sparkline's data range
            List<double> sparklineAverages = new List<double>();

            // Iterate through each sparkline in all groups
            foreach (SparklineGroup group in worksheet.SparklineGroups)
            {
                foreach (Sparkline sparkline in group.Sparklines)
                {
                    // Obtain the range referenced by the sparkline
                    Aspose.Cells.Range dataRange = worksheet.Cells.CreateRange(sparkline.DataRange);

                    double sum = 0;
                    int count = 0;

                    // Accumulate numeric values
                    foreach (Cell cell in dataRange)
                    {
                        if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
                        {
                            sum += val;
                            count++;
                        }
                    }

                    // Compute average (avoid division by zero)
                    double average = count > 0 ? sum / count : 0;
                    sparklineAverages.Add(average);
                }
            }

            // Output the collected averages
            Console.WriteLine("Average values of each sparkline:");
            foreach (double avg in sparklineAverages)
            {
                Console.WriteLine(avg);
            }

            // Save the workbook (ensure directory exists)
            string outputPath = "SparklinesWithAverages.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}