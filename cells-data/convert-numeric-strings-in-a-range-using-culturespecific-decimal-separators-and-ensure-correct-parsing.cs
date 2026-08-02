using System;
using Aspose.Cells;

class ConvertNumericStringsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set culture‑specific separators (comma as decimal, dot as group)
        wb.Settings.NumberDecimalSeparator = ',';
        wb.Settings.NumberGroupSeparator = '.';

        // Get the cells collection of the first worksheet
        Cells cells = wb.Worksheets[0].Cells;

        // Populate a range with numeric strings that use the comma decimal separator
        cells["A1"].PutValue("1,23");
        cells["A2"].PutValue("4,567");
        cells["A3"].PutValue("12,34");
        cells["A4"].PutValue("7,89");
        cells["A5"].PutValue("0,001");

        // Populate another range with strings that use a dot decimal separator (will stay as strings)
        cells["B1"].PutValue("5.67");
        cells["B2"].PutValue("8.90");
        cells["B3"].PutValue("3.1415");
        cells["B4"].PutValue("2.718");
        cells["B5"].PutValue("6.022e23");

        // Convert string values to numeric where possible, respecting the workbook's separators
        cells.ConvertStringToNumericValue();

        // Display the conversion results
        for (int row = 0; row < 5; row++)
        {
            Cell cellA = cells[row, 0]; // Column A
            Cell cellB = cells[row, 1]; // Column B

            string valueA = cellA.Type == CellValueType.IsNumeric
                ? cellA.DoubleValue.ToString()
                : cellA.StringValue;

            string valueB = cellB.Type == CellValueType.IsNumeric
                ? cellB.DoubleValue.ToString()
                : cellB.StringValue;

            Console.WriteLine($"A{row + 1}: {valueA}   B{row + 1}: {valueB}");
        }

        // Save the workbook to verify the result in Excel
        wb.Save("ConvertedNumbers.xlsx");
    }
}