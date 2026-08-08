// Title: Aspose.Cells C# – Enumerate Worksheet Columns and Get Maximum Numeric Value per Column
// Description: This example creates a workbook, fills three columns with numbers, obtains a non‑generic Columns enumerator, walks each column, scans rows up to the last data row, detects numeric cells, records the highest value per column in a dictionary, prints the column name with its max, and saves the file.
// Keywords: Aspose.Cells column enumerator | C# iterate columns Excel | maximum numeric value per column | Columns.GetEnumerator Aspose | CellsHelper.ColumnIndexToName | Excel column max value .NET | Aspose.Cells data analysis example | non‑generic enumerator C#
// Common Searches: How to enumerate columns with Aspose.Cells in C# | Get max numeric cell value for each Excel column using Aspose.Cells | Aspose.Cells C# column-wise maximum calculation | Iterate worksheet columns and find highest value Aspose.Cells
// Developer Intent: Retrieve the highest numeric cell value for every column in a worksheet using Aspose.Cells.
// Use Cases: Generate a summary table that lists the peak value of each column for financial reporting. | Validate data ranges by ensuring column maxima stay within acceptable thresholds. | Extract column‑wise maxima to feed a chart or dashboard that highlights extreme measurements.
// AI Prompts: Write C# code with Aspose.Cells that enumerates all columns and returns a Dictionary<int, double> of column indexes and their maximum numeric values. | Provide a LINQ‑based solution to compute the maximum numeric value per column in an Aspose.Cells worksheet. | Explain how to safely handle columns that contain no numeric cells when calculating column maxima with Aspose.Cells.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsColumnMaxValueDemo
{
    // This example creates a workbook, fills three columns with numbers, obtains a non‑generic Columns enumerator, walks each column, scans rows up to the last data row, detects numeric cells, records the highest value per column in a dictionary, prints the column name with its max, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample numeric data across several columns
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(25);
                cells["A3"].PutValue(5);

                cells["B1"].PutValue(7);
                cells["B2"].PutValue(14);
                cells["B3"].PutValue(21);

                cells["C1"].PutValue(30);
                cells["C2"].PutValue(12);
                cells["C3"].PutValue(18);

                // Dictionary to hold the maximum numeric value per column (key = column index)
                Dictionary<int, double> maxValuesPerColumn = new Dictionary<int, double>();

                // Iterate each column using the non‑generic enumerator
                IEnumerator columnEnumerator = cells.Columns.GetEnumerator();
                while (columnEnumerator.MoveNext())
                {
                    // Current column object
                    Column column = (Column)columnEnumerator.Current;

                    // Determine the column index (zero‑based)
                    int colIndex = cells.Columns.IndexOf(column);
                    if (colIndex < 0) continue; // safety check

                    double maxInColumn = double.MinValue;
                    bool hasNumeric = false;

                    // Determine the last row that may contain data to limit the loop
                    int lastRow = cells.MaxDataRow;
                    if (lastRow < 0) lastRow = 0; // no data case

                    // Iterate through each row in the current column
                    for (int row = 0; row <= lastRow; row++)
                    {
                        Cell cell = cells[row, colIndex];
                        if (cell != null && cell.Value != null && cell.Type == CellValueType.IsNumeric)
                        {
                            double val = cell.DoubleValue;
                            if (!hasNumeric || val > maxInColumn)
                            {
                                maxInColumn = val;
                                hasNumeric = true;
                            }
                        }
                    }

                    // Record the maximum value if any numeric cell was found
                    if (hasNumeric)
                    {
                        maxValuesPerColumn[colIndex] = maxInColumn;
                    }
                }

                // Output the results
                Console.WriteLine("Maximum numeric value per column:");
                foreach (var kvp in maxValuesPerColumn)
                {
                    // Convert column index to Excel column name for readability
                    string columnName = CellsHelper.ColumnIndexToName(kvp.Key);
                    Console.WriteLine($"{columnName}: {kvp.Value}");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "ColumnMaxValuesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
