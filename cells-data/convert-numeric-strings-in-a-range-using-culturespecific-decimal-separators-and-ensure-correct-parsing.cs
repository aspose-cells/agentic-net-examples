using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Configure culture‑specific separators: comma as decimal, dot as group separator
        workbook.Settings.NumberDecimalSeparator = ',';
        workbook.Settings.NumberGroupSeparator = '.';

        // Get the first worksheet's cells collection
        Cells cells = workbook.Worksheets[0].Cells;

        // Insert numeric strings that use the configured separators
        cells["A1"].PutValue("123,45");      // Simple decimal number
        cells["A2"].PutValue("1.234,56");    // Number with group separator
        cells["A3"].PutValue("not a number"); // This will remain a string

        // Convert all convertible string values in the worksheet to numeric values
        cells.ConvertStringToNumericValue(); // feature rule: ConvertStringToNumericValue

        // Output the results to verify correct parsing
        Console.WriteLine("A1 (numeric): " + cells["A1"].DoubleValue);
        Console.WriteLine("A2 (numeric): " + cells["A2"].DoubleValue);
        Console.WriteLine("A3 (string) : " + cells["A3"].StringValue);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("ConvertedNumbers.xlsx");
    }
}