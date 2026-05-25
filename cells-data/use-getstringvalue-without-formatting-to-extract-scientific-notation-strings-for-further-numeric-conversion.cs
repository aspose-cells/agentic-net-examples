using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsScientificNotationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a scientific notation string into a cell (as text)
            cells["A1"].PutValue("1.23E+5");

            // Retrieve the raw string without any formatting
            string rawString = cells["A1"].GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine("Raw string from cell: " + rawString);

            // Convert the scientific notation string to a numeric value
            double numericValue = double.Parse(rawString, CultureInfo.InvariantCulture);
            Console.WriteLine("Converted numeric value: " + numericValue);

            // Optionally, write the numeric value back to another cell
            cells["B1"].PutValue(numericValue);
            Console.WriteLine("Numeric value written to B1: " + cells["B1"].DoubleValue);

            // Save the workbook (demonstrates lifecycle usage)
            workbook.Save("ScientificNotationDemo.xlsx");
        }
    }
}