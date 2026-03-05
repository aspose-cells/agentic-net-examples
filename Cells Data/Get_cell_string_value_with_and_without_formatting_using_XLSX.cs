using System;
using Aspose.Cells;

namespace AsposeCellsStringValueDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put different types of data into cells
            cells["A1"].PutValue("Plain Text");
            cells["A2"].PutValue(12345);               // Numeric
            cells["A3"].PutValue(3.14159);             // Double
            cells["A4"].PutValue(DateTime.Now);        // DateTime

            // Apply number formatting to demonstrate formatted output
            Style style = cells["A2"].GetStyle();
            style.Number = 3; // Currency format
            cells["A2"].SetStyle(style);

            style = cells["A3"].GetStyle();
            style.Number = 9; // Percentage format
            cells["A3"].SetStyle(style);

            // Retrieve values with formatting (StringValue)
            Console.WriteLine("=== Formatted StringValue ===");
            Console.WriteLine($"A1: {cells["A1"].StringValue}");
            Console.WriteLine($"A2: {cells["A2"].StringValue}");
            Console.WriteLine($"A3: {cells["A3"].StringValue}");
            Console.WriteLine($"A4: {cells["A4"].StringValue}");

            // Retrieve values without formatting using GetStringValue with CellValueFormatStrategy.None
            Console.WriteLine("\n=== Unformatted StringValue (CellValueFormatStrategy.None) ===");
            Console.WriteLine($"A1: {cells["A1"].GetStringValue(CellValueFormatStrategy.None)}");
            Console.WriteLine($"A2: {cells["A2"].GetStringValue(CellValueFormatStrategy.None)}");
            Console.WriteLine($"A3: {cells["A3"].GetStringValue(CellValueFormatStrategy.None)}");
            Console.WriteLine($"A4: {cells["A4"].GetStringValue(CellValueFormatStrategy.None)}");

            // Save the workbook to an XLSX file
            workbook.Save("StringValueDemo.xlsx");
        }
    }
}