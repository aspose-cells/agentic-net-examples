// Title: C# – Compute and Collect Average Values of Sparkline Data with Aspose.Cells
// Description: Creates a workbook, fills a 3×4 range, adds line sparklines per row, then iterates every sparkline group, extracts each sparkline's data range, calculates the numeric average of the cells, stores the results in a List<double>, prints the averages, and saves the file.
// Keywords: Aspose.Cells | C# | sparkline average | iterate sparklines | sparkline data range | calculate sparkline statistics | Excel sparkline processing | Aspose.Cells API
// Common Searches: Aspose.Cells calculate sparkline average C# | How to get sparkline data values with Aspose.Cells | Iterate sparkline groups .NET | Store sparkline averages in a list Aspose.Cells | Retrieve sparkline range values C#
// Developer Intent: Extract each sparkline’s data range, compute its average, and collect the results in a list.
// Use Cases: Build a summary sheet that lists the average value of each row’s sparkline for quick trend analysis. | Apply conditional formatting based on sparkline averages to highlight high‑ or low‑performing rows. | Export the calculated averages to another worksheet or a CSV file for downstream reporting.
// AI Prompts: Generate a reusable method that receives a Worksheet and returns a List<double> of all sparkline averages using Aspose.Cells. | Provide error‑handling code that safely computes sparkline averages when the range contains non‑numeric or empty cells. | Show how to write each computed sparkline average back to the worksheet next to the sparkline column.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills a 3×4 range, adds line sparklines per row, then iterates every sparkline group, extracts each sparkline's data range, calculates the numeric average of the cells, stores the results in a List<double>, prints the averages, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (3 rows x 4 columns)
            for (int r = 0; r < 3; r++)
            {
                for (int c = 0; c < 4; c++)
                {
                    worksheet.Cells[r, c].PutValue((r + 1) * (c + 1));
                }
            }

            // Define where the sparklines will be placed (column E)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group for the data range A1:D3
            int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D3", false, location);
            SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

            // Add a sparkline for each row in the data range
            for (int row = 0; row < 3; row++)
            {
                // Data range for the current row, e.g., "A1:D1"
                string dataRange = $"{worksheet.Name}!A{row + 1}:D{row + 1}";
                sparklineGroup.Sparklines.Add(dataRange, row, 4);
            }

            // List to store average values of each sparkline
            List<double> sparklineAverages = new List<double>();

            // Iterate through all sparkline groups and their sparklines
            foreach (SparklineGroup group in worksheet.SparklineGroups)
            {
                foreach (Sparkline sparkline in group.Sparklines)
                {
                    // Get the data range string (may include sheet name)
                    string range = sparkline.DataRange;
                    // Remove sheet name if present
                    string address = range.Contains("!") ? range.Split('!')[1] : range;

                    // Retrieve the range object from the worksheet
                    Aspose.Cells.Range cellsRange = worksheet.Cells.CreateRange(address);

                    double sum = 0;
                    int count = 0;

                    // Calculate sum and count of numeric values in the range
                    foreach (Cell cell in cellsRange)
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

            // Output the averages to console
            for (int i = 0; i < sparklineAverages.Count; i++)
            {
                Console.WriteLine($"Sparkline {i} average = {sparklineAverages[i]}");
            }

            // Save the workbook
            workbook.Save("SparklinesWithAverages.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
