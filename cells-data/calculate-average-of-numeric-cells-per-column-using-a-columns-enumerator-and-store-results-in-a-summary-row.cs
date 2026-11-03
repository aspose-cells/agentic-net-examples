// Title: Compute column averages of numeric cells and add a summary row with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells calculate average for each column and write to a new row | How to add a summary row with column averages using Aspose.Cells in .NET | Iterate through columns with MaxDataColumn to compute numeric averages in Aspose.Cells | Write average values to the row after data range in an Excel workbook using Aspose.Cells C# | Determine last data row and column then insert average row with Aspose.Cells
// Tags: column average computation Aspose.Cells | insert average row Excel Aspose.Cells | enumerate columns using MaxDataColumn Aspose.Cells | numeric cell detection Aspose.Cells | write values to worksheet Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, fills a 5x4 range with numeric data, finds the last occupied row and column, adds a label row, iterates each column to sum numeric cells, computes the average, writes the averages into a new summary row, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data (rows 0-4, columns 0-3)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue(row * 10 + col + 1); // example values
                }
            }

            // Determine the last row and column that contain data
            int lastDataRow = sheet.Cells.MaxDataRow;
            int lastDataColumn = sheet.Cells.MaxDataColumn;

            // Row where the summary (averages) will be written
            int summaryRow = lastDataRow + 1;

            // Write a label for the summary row
            sheet.Cells[summaryRow, 0].PutValue("Average");

            // Iterate through each column that has data
            for (int col = 0; col <= lastDataColumn; col++)
            {
                double sum = 0;
                int count = 0;

                // Iterate through rows up to the last data row for the current column
                for (int row = 0; row <= lastDataRow; row++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        sum += cell.DoubleValue;
                        count++;
                    }
                }

                // Calculate and write the average if numeric cells were found
                if (count > 0)
                {
                    double average = sum / count;
                    sheet.Cells[summaryRow, col].PutValue(average);
                }
            }

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
