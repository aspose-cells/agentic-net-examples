// Title: Find the maximum numeric value in a worksheet by enumerating columns with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to loop through every column in a worksheet and return the highest numeric cell value found. | Extend the column‑scanning logic to also consider decimal, float, and long types when calculating the maximum numeric value. | Add safeguards to ignore empty, null, or non‑numeric cells while aggregating the overall maximum across all columns.
// Common Searches: asp.net find highest numeric cell value in Excel using Aspose.Cells column enumeration | c# Aspose.Cells iterate columns to get max number in worksheet | how to calculate overall maximum numeric value across columns with Aspose.Cells for .NET | retrieve maximum double or int from Excel sheet using Aspose.Cells Columns collection
// Tags: enumerate columns Aspose.Cells C# | maximum numeric cell value Aspose.Cells | column-wise scanning Aspose.Cells | numeric type handling Aspose.Cells | worksheet data range Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsColumnMaxValueDemo
{
    // The program creates a workbook, fills cells with numeric and non‑numeric data, enumerates each column via Cells.Columns, iterates rows between MinDataRow and MaxDataRow, checks for int and double values (with placeholders for other numeric types), tracks the highest numeric value, outputs the result, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample numeric and non‑numeric data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(25);
            cells["A3"].PutValue("Text");
            cells["B1"].PutValue(5);
            cells["B2"].PutValue(30);
            cells["C1"].PutValue(12.5);
            cells["C2"].PutValue(null);
            cells["D1"].PutValue(100);
            cells["D2"].PutValue("Another text");

            // Variable to hold the overall maximum numeric value
            double maxNumericValue = double.MinValue;
            bool foundNumeric = false;

            // Obtain the Columns enumerator and iterate each column
            foreach (Column column in cells.Columns)
            {
                // The Column class provides the Index property (zero‑based column index)
                int colIndex = column.Index;

                // Determine the range of rows that may contain data
                int startRow = cells.MinDataRow;
                int endRow = cells.MaxDataRow;

                // If there are no data rows, skip this column
                if (startRow == -1 || endRow == -1)
                    continue;

                // Iterate through each row in the current column
                for (int row = startRow; row <= endRow; row++)
                {
                    Cell cell = cells[row, colIndex];
                    if (cell == null || cell.Value == null)
                        continue;

                    // Check if the cell contains a numeric value
                    if (cell.Value is double d)
                    {
                        if (!foundNumeric || d > maxNumericValue)
                        {
                            maxNumericValue = d;
                            foundNumeric = true;
                        }
                    }
                    else if (cell.Value is int i)
                    {
                        double dVal = i;
                        if (!foundNumeric || dVal > maxNumericValue)
                        {
                            maxNumericValue = dVal;
                            foundNumeric = true;
                        }
                    }
                    // Add other numeric types if needed (e.g., decimal, float)
                }
            }

            // Output the result
            if (foundNumeric)
                Console.WriteLine("Maximum numeric cell value: " + maxNumericValue);
            else
                Console.WriteLine("No numeric cells were found.");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ColumnMaxValueDemo.xlsx");
        }
    }
}
