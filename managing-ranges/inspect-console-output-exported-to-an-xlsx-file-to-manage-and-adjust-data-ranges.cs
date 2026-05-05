using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Load the workbook that was exported from console output
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Retrieve the maximal display range (covers data, merged cells, shapes)
        AsposeRange maxRange = cells.MaxDisplayRange;

        if (maxRange != null)
        {
            // Output current range information to the console
            Console.WriteLine($"Current data range: Rows {maxRange.FirstRow} - {maxRange.FirstRow + maxRange.RowCount - 1}, " +
                              $"Columns {maxRange.FirstColumn} - {maxRange.FirstColumn + maxRange.ColumnCount - 1}");

            // Example adjustment: insert a new column right after the existing data range
            int newColumnIndex = maxRange.FirstColumn + maxRange.ColumnCount;
            cells.InsertColumn(newColumnIndex);

            // Add a header for the new column
            cells[maxRange.FirstRow, newColumnIndex].PutValue("NewColumn");

            // Populate the new column with sequential numbers (skip header row)
            for (int r = 1; r < maxRange.RowCount; r++)
            {
                cells[maxRange.FirstRow + r, newColumnIndex].PutValue(r);
            }
        }
        else
        {
            Console.WriteLine("Worksheet is empty; no range to adjust.");
        }

        // Remove any blank rows or columns that may have been created
        cells.DeleteBlankRows();
        cells.DeleteBlankColumns();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}