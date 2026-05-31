using System;
using Aspose.Cells;

namespace FreezeFirstTwoColumnsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the first two columns (A and B)
            // Freeze at column index 2 (C) with 0 frozen rows and 2 frozen columns
            worksheet.FreezePanes(0, 2, 0, 2);

            // Save the workbook
            workbook.Save("FreezeFirstTwoColumns.xlsx");
        }
    }
}