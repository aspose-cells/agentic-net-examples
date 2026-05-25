using System;
using Aspose.Cells;

namespace AsposeCellsStringValueDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access a specific cell (A1) and put a numeric value into it
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue(12345.678);

            // Apply a number format (e.g., currency) to the cell
            Style style = cell.GetStyle();
            style.Number = 3; // Currency format
            cell.SetStyle(style);

            // Retrieve the cell's formatted string value using the DisplayString strategy
            // This strategy includes all formatting (number format, locale, column width, etc.)
            string formattedValue = cell.GetStringValue(CellValueFormatStrategy.DisplayString);

            // Output the result
            Console.WriteLine("Formatted string value (DisplayString strategy): " + formattedValue);
        }
    }
}