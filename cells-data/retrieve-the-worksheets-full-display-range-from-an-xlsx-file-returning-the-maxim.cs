using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        Aspose.Cells.Range maxDisplayRange = worksheet.Cells.MaxDisplayRange;

        if (maxDisplayRange != null)
        {
            Console.WriteLine("Max Display Range:");
            Console.WriteLine($"Start Row (0‑based): {maxDisplayRange.FirstRow}");
            Console.WriteLine($"Start Column (0‑based): {maxDisplayRange.FirstColumn}");
            Console.WriteLine($"Total Rows: {maxDisplayRange.RowCount}");
            Console.WriteLine($"Total Columns: {maxDisplayRange.ColumnCount}");
            Console.WriteLine($"Address: {maxDisplayRange.Address}");
        }
        else
        {
            Console.WriteLine("The worksheet is empty; no display range found.");
        }

        workbook.Save("output.xlsx");
    }
}