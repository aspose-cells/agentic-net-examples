using System;
using Aspose.Cells;

namespace AsposeCellsGetStringValueDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put a numeric value into cell A1
            cells["A1"].PutValue(12345.6789);

            // Apply a number format to the cell (e.g., currency)
            Style style = cells["A1"].GetStyle();
            style.Number = 4; // Currency format
            cells["A1"].SetStyle(style);

            // Retrieve the formatted string using the DisplayString strategy
            string formattedValue = cells["A1"].GetStringValue(CellValueFormatStrategy.DisplayString);
            Console.WriteLine("Formatted (DisplayString) value: " + formattedValue);

            // Retrieve the raw value without any formatting using the None strategy
            string rawValue = cells["A1"].GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine("Raw (None) value: " + rawValue);

            // Optionally save the workbook (not required for the GetStringValue demonstration)
            workbook.Save("GetStringValueDemo.xlsx");
        }
    }
}