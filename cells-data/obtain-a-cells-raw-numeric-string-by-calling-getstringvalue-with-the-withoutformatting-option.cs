using System;
using Aspose.Cells;

namespace AsposeCellsRawStringDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put a numeric value into cell A1
            cells["A1"].PutValue(12345.6789);

            // Obtain the raw numeric string without any formatting
            // CellValueFormatStrategy.None means no formatting is applied
            string rawNumericString = cells["A1"].GetStringValue(CellValueFormatStrategy.None);

            // Display the result
            Console.WriteLine("Raw numeric string (no formatting): " + rawNumericString);

            // Save the workbook (optional)
            workbook.Save("RawStringDemo.xlsx");
        }
    }
}