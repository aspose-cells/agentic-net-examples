// Title: How to get an unformatted numeric string from a cell using Cell.GetStringValue with CellValueFormatStrategy.None in Aspose.Cells for .NET
// AI Prompts: Generate C# code that calls Cell.GetStringValue with CellValueFormatStrategy.None to return the exact numeric text stored in a worksheet cell, ignoring any applied number format. | Show how to extract the raw numeric string from a formatted Excel cell using Aspose.Cells without triggering any formatting conversion.
// Common Searches: Aspose.Cells C# get cell value as plain string ignoring number format | Cell.GetStringValue CellValueFormatStrategy.None example code | Retrieve raw numeric text from formatted Excel cell using Aspose.Cells .NET | How to read unformatted numeric string from Excel with Aspose.Cells in C# | Get underlying numeric string from a currency-formatted cell Aspose.Cells
// Tags: Aspose.Cells GetStringValue raw numeric string | CellValueFormatStrategy.None unformatted cell value | C# extract numeric string from formatted Excel cell | Aspose.Cells retrieve cell text without formatting | Excel numeric value as string Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a workbook, writes a numeric value to cell A1, applies a currency format, then uses Cell.GetStringValue with CellValueFormatStrategy.None to obtain the raw numeric string without any formatting and prints it.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a numeric value into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue(12345.6789);

        // Apply a number format (e.g., currency) to demonstrate the difference
        Style style = cell.GetStyle();
        style.Number = 44; // Currency format
        cell.SetStyle(style);

        // Retrieve the raw numeric string without any formatting
        string rawNumericString = cell.GetStringValue(CellValueFormatStrategy.None);
        Console.WriteLine("Raw numeric string (no formatting): " + rawNumericString);

        // Save the workbook (optional)
        workbook.Save("RawNumericStringDemo.xlsx");
    }
}
