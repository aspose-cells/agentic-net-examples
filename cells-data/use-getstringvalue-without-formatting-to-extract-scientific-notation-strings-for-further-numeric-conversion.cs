using System;
using Aspose.Cells;
using System.Globalization;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a value that is stored as a scientific notation string
        worksheet.Cells["A1"].PutValue("1.23E+04");

        // Retrieve the cell's raw string without applying any formatting
        // CellValueFormatStrategy.None returns the unformatted string representation
        string rawString = worksheet.Cells["A1"].GetStringValue(CellValueFormatStrategy.None);
        Console.WriteLine("Raw string from cell: " + rawString);

        // Convert the scientific notation string to a numeric double value
        if (double.TryParse(rawString, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue))
        {
            Console.WriteLine("Converted numeric value: " + numericValue);
        }
        else
        {
            Console.WriteLine("Failed to convert the string to a numeric value.");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("ScientificNotationDemo.xlsx");
    }
}