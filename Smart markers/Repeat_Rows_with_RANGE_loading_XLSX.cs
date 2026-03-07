using System;
using Aspose.Cells;

class RepeatRowsExample
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Specify the rows that should be repeated on every printed page.
        // Here rows 1 and 2 (Excel rows are 1‑based) will be repeated.
        worksheet.PageSetup.PrintTitleRows = "$1:$2";

        // Save the workbook with the updated print title settings
        workbook.Save("output.xlsx");
    }
}