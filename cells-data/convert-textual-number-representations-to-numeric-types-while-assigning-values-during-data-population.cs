using System;
using Aspose.Cells;

namespace AsposeCellsNumberConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with textual representations of numbers.
            // The second parameter 'true' enables conversion to the appropriate data type.
            cells["A1"].PutValue("123.45", true);          // Converted to double
            cells["A2"].PutValue("678", true);            // Converted to int (stored as double internally)
            cells["A3"].PutValue("2023-07-08", true);     // Converted to DateTime
            cells["A4"].PutValue("NotANumber", true);     // Remains as string because conversion fails

            // Additionally, put some values without conversion flag to demonstrate later conversion.
            cells["B1"].PutValue("987.65");               // Stored as string initially
            cells["B2"].PutValue("2021/12/31");           // Stored as string initially

            // Convert any remaining string values that can be interpreted as numeric or date.
            // This uses the Cells.ConvertStringToNumericValue method.
            cells.ConvertStringToNumericValue();

            // Display the resulting types and values in the console.
            Console.WriteLine($"A1: {cells["A1"].Value} (Type: {cells["A1"].Type})");
            Console.WriteLine($"A2: {cells["A2"].Value} (Type: {cells["A2"].Type})");
            Console.WriteLine($"A3: {cells["A3"].Value} (Type: {cells["A3"].Type})");
            Console.WriteLine($"A4: {cells["A4"].Value} (Type: {cells["A4"].Type})");
            Console.WriteLine($"B1: {cells["B1"].Value} (Type: {cells["B1"].Type})");
            Console.WriteLine($"B2: {cells["B2"].Value} (Type: {cells["B2"].Type})");

            // Save the workbook to an Excel file (lifecycle save rule)
            workbook.Save("NumberConversionResult.xlsx");
        }
    }
}