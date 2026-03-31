using System;
using Aspose.Cells;

namespace AsposeCellsZeroDisplayDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Suppress visual display of zero-valued cells while keeping the data
            sheet.DisplayZeros = false;

            // Add some sample data, including zero values
            sheet.Cells["A1"].PutValue(123);   // non-zero value
            sheet.Cells["A2"].PutValue(0);     // zero value (will not be displayed)
            sheet.Cells["A3"].PutValue(456);   // non-zero value
            sheet.Cells["A4"].PutValue(0);     // zero value (will not be displayed)

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ZeroDisplaySuppressed.xlsx", SaveFormat.Xlsx);
        }
    }
}