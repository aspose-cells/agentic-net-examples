using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze the first row (header) so it remains visible while scrolling
        // Parameters: row index (1), column index (0), frozen rows (1), frozen columns (0)
        worksheet.FreezePanes(1, 0, 1, 0);

        // Save the modified workbook to a new XLSX file
        workbook.Save("FrozenHeader.xlsx");
    }
}