using System;
using Aspose.Cells;

class ShowRowsAndColumnsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (5 rows x 4 columns)
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Set custom height for the third row (index 2) and width for the second column (index 1)
        worksheet.Cells.Rows[2].Height = 30;          // Row height in points
        worksheet.Cells.SetColumnWidth(1, 25);        // Column width in characters

        // ---------- Show Rows ----------
        Console.WriteLine("Rows:");
        // Iterate through the RowCollection
        foreach (Row row in worksheet.Cells.Rows)
        {
            // Row.Index gives the zero‑based row number
            Console.WriteLine($"Row {row.Index} - Height: {row.Height}");
        }

        // ---------- Show Columns ----------
        Console.WriteLine("\nColumns:");
        // Assuming we have 4 columns (0 to 3) based on the data inserted above
        for (int col = 0; col < 4; col++)
        {
            // GetColumnWidth returns the width in characters for normal view
            double width = worksheet.Cells.GetColumnWidth(col);
            Console.WriteLine($"Column {col} - Width: {width}");
        }

        // Save the workbook to demonstrate the changes
        workbook.Save("ShowRowsAndColumns.xlsx");
    }
}