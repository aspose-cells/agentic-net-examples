// Title: Get a formatted cell string using GetStringValue with CellValueFormatStrategy.WithFormatting in Aspose.Cells for .NET
// Description: Creates a workbook, sets a numeric value in A1, applies a currency number format, and demonstrates how to obtain the cell's display text with GetStringValue using the WithFormatting strategy (and compares it to DisplayString and CellStyle). The formatted string is printed to the console and the workbook can be saved.
// Keywords: Aspose.Cells GetStringValue WithFormatting | formatted cell value C# | CellValueFormatStrategy.WithFormatting example | currency format Aspose.Cells | retrieve display string from cell
// Common Searches: Aspose.Cells GetStringValue WithFormatting | how to get formatted cell text in C# Aspose | CellValueFormatStrategy.WithFormatting sample code | retrieve currency formatted value Aspose.Cells | display string of a cell using Aspose.Cells
// Developer Intent: Obtain the exact display string of a cell that reflects its applied number format.
// Use Cases: Show a currency‑formatted amount directly in a UI without exporting the workbook. | Log cell values with their visual formatting for audit trails. | Compare different GetStringValue strategies to choose the appropriate representation for reporting.
// AI Prompts: Write C# code that uses Aspose.Cells GetStringValue with CellValueFormatStrategy.WithFormatting to return a formatted string from a cell. | Explain the differences between DisplayString, CellStyle, and WithFormatting strategies in GetStringValue. | Demonstrate how to retrieve a date‑formatted cell value as a string using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, sets a numeric value in A1, applies a currency number format, and demonstrates how to obtain the cell's display text with GetStringValue using the WithFormatting strategy (and compares it to DisplayString and CellStyle). The formatted string is printed to the console and the workbook can be saved.
class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a numeric value into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue(12345.6789);

        // Apply a currency number format to the cell
        Style style = cell.GetStyle();
        style.Number = 4; // Currency format
        cell.SetStyle(style);

        // Retrieve the formatted string using GetStringValue with the DisplayString strategy
        string formattedValue = cell.GetStringValue(CellValueFormatStrategy.DisplayString);
        Console.WriteLine("Formatted (DisplayString): " + formattedValue);

        // Retrieve the formatted string using the CellStyle strategy for comparison
        string cellStyleValue = cell.GetStringValue(CellValueFormatStrategy.CellStyle);
        Console.WriteLine("Formatted (CellStyle): " + cellStyleValue);

        // Save the workbook (optional)
        workbook.Save("FormattedCell.xlsx");
    }
}
