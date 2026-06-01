using System;
using System.Collections;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample numeric data in the first column (A)
        for (int i = 0; i < 10; i++)
        {
            // Cells[row, column] – column index 0 corresponds to column A
            worksheet.Cells[i, 0].PutValue(i + 1); // values 1,2,...,10
        }

        // Obtain an enumerator for the rows collection
        IEnumerator rowsEnumerator = worksheet.Cells.Rows.GetEnumerator();

        double sum = 0;

        // Iterate through each row
        while (rowsEnumerator.MoveNext())
        {
            Row row = (Row)rowsEnumerator.Current;

            // Access the cell in the first column of the current row
            Cell cell = row[0]; // equivalent to worksheet.Cells[row.Index, 0]

            // Ensure the cell contains a numeric value before adding
            if (cell != null && cell.Value != null && double.TryParse(cell.Value.ToString(), out double value))
            {
                sum += value;
            }
        }

        Console.WriteLine($"Sum of values in the first column: {sum}");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("RowsSumDemo.xlsx");
    }
}