using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsColumnMaxValueDemo
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
                Cells cells = worksheet.Cells;

                // Populate sample data (numeric and non‑numeric) in several columns
                cells["A1"].PutValue("HeaderA");
                cells["A2"].PutValue(10);
                cells["A3"].PutValue(25);
                cells["A4"].PutValue(7);

                cells["B1"].PutValue("HeaderB");
                cells["B2"].PutValue(5);
                cells["B3"].PutValue(15);
                cells["B4"].PutValue("Text");   // non‑numeric

                cells["C1"].PutValue("HeaderC");
                cells["C2"].PutValue("String"); // non‑numeric
                cells["C3"].PutValue(30);
                cells["C4"].PutValue(20);

                // Dictionary to store the maximum numeric value per column (key = column index)
                Dictionary<int, double> maxValuesPerColumn = new Dictionary<int, double>();

                // Obtain an enumerator for the Columns collection
                IEnumerator columnEnumerator = worksheet.Cells.Columns.GetEnumerator();

                // Iterate each column
                while (columnEnumerator.MoveNext())
                {
                    // Cast the current item to Column
                    Column column = (Column)columnEnumerator.Current;

                    // Column index (0‑based)
                    int colIndex = column.Index;

                    double maxInColumn = double.MinValue;
                    bool hasNumeric = false;

                    // Determine the last row that contains data to limit the scan
                    int lastRow = cells.MaxDataRow;
                    if (lastRow < 0) lastRow = -1; // no data at all

                    // Scan all rows in the current column
                    for (int row = 0; row <= lastRow; row++)
                    {
                        Cell cell = cells[row, colIndex];
                        if (cell == null || cell.Value == null)
                            continue;

                        // Check for numeric types and update the maximum
                        if (cell.Value is double d)
                        {
                            hasNumeric = true;
                            if (d > maxInColumn) maxInColumn = d;
                        }
                        else if (cell.Value is int i)
                        {
                            hasNumeric = true;
                            double dVal = i;
                            if (dVal > maxInColumn) maxInColumn = dVal;
                        }
                        else if (cell.Value is decimal dec)
                        {
                            hasNumeric = true;
                            double dVal = (double)dec;
                            if (dVal > maxInColumn) maxInColumn = dVal;
                        }
                    }

                    // Record the maximum if at least one numeric value was found
                    if (hasNumeric)
                    {
                        maxValuesPerColumn[colIndex] = maxInColumn;
                    }
                }

                // Output the results
                foreach (var kvp in maxValuesPerColumn)
                {
                    Console.WriteLine($"Column {kvp.Key} (\"{cells[0, kvp.Key].StringValue}\") max numeric value: {kvp.Value}");
                }

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                string outputPath = "ColumnsMaxValues.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}