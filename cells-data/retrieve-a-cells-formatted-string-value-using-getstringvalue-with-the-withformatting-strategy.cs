// Title: Retrieve a formatted string from an Excel cell using GetStringValue with DisplayString strategy in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that writes a numeric value to cell A1, applies a currency number format, and reads the formatted text using Aspose.Cells GetStringValue with CellValueFormatStrategy.DisplayString. | Show how to obtain both the formatted and raw string representations of a cell value in Aspose.Cells by calling GetStringValue with DisplayString and None strategies. | Provide a complete .NET example that saves the workbook after extracting formatted cell strings with GetStringValue.
// Common Searches: Aspose.Cells C# GetStringValue DisplayString example for formatted currency | How to read formatted cell value as string using Aspose.Cells .NET | CellValueFormatStrategy.DisplayString vs None in Aspose.Cells | Retrieve formatted Excel cell text with GetStringValue in C# | Aspose.Cells get string value with formatting strategy
// Tags: Aspose.Cells GetStringValue formatted string C# | CellValueFormatStrategy DisplayString usage | apply currency number format Aspose.Cells | retrieve raw cell string Aspose.Cells | save workbook after GetStringValue Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsGetStringValueDemo
{
    // Creates a workbook, inserts a numeric value into A1, applies a currency format, then uses GetStringValue with CellValueFormatStrategy.DisplayString to obtain the formatted string and with CellValueFormatStrategy.None to obtain the raw value, finally saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a numeric value into cell A1
            cells["A1"].PutValue(12345.6789);

            // Apply a number format (currency) to the cell
            Style style = cells["A1"].GetStyle();
            style.Number = 4; // Currency format
            cells["A1"].SetStyle(style);

            // Retrieve the formatted string using the DisplayString strategy
            string formattedValue = cells["A1"].GetStringValue(CellValueFormatStrategy.DisplayString);

            // Output the result
            Console.WriteLine("Formatted string (DisplayString strategy): " + formattedValue);

            // Optionally, retrieve the raw value without formatting
            string rawValue = cells["A1"].GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine("Raw string (None strategy): " + rawValue);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("GetStringValueDemo.xlsx");
        }
    }
}
