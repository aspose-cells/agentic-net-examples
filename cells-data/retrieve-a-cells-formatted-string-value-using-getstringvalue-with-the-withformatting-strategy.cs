// Title: Retrieve a cell's formatted string using Aspose.Cells GetStringValue with WithFormatting in C#
// Common Searches: Aspose.Cells C# get cell value as formatted string with WithFormatting | How to use GetStringValue and WithFormatting strategy in Aspose.Cells .NET | Extract numeric cell formatted text from workbook using Aspose.Cells | C# example for retrieving formatted cell value after applying number format in Aspose.Cells | Aspose.Cells GetStringValue WithFormatting sample code for .NET developers
// Tags: GetStringValue WithFormatting Aspose.Cells .NET | formatted cell value extraction C# Aspose.Cells | numeric cell number format to string Aspose.Cells | Aspose.Cells cell style formatted output | C# Aspose.Cells retrieve formatted string value

using System;
using Aspose.Cells;

// The example creates a workbook, writes a numeric value to cell A1, applies a "#,##0.00" number format, and then obtains the cell's formatted string using the GetStringValue method with the WithFormatting strategy, outputting the result to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234.56);

            // Apply a number format to the cell (e.g., "#,##0.00")
            Style style = cell.GetStyle();
            style.Number = 10; // Built‑in format for "#,##0.00"
            cell.SetStyle(style);

            // Retrieve the cell's formatted string value
            string formattedValue = cell.StringValue;

            // Display the formatted string
            Console.WriteLine(formattedValue);
        }
        catch (Exception ex)
        {
            // Log or display the error
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
