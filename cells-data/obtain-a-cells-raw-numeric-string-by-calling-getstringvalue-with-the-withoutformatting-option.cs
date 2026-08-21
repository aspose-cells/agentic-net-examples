// Title: Get raw numeric string from a cell with GetStringValue(CellValueFormatStrategy.None) – Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, writes the number 12345.6789 to cell A1, then uses GetStringValue(CellValueFormatStrategy.None) to obtain the exact numeric text without any formatting, displays it, and saves the workbook.
// Keywords: Aspose.Cells | GetStringValue | CellValueFormatStrategy.None | raw numeric string | C# .NET | unformatted cell value | Excel numeric extraction
// Common Searches: Aspose.Cells GetStringValue without formatting | How to read original numeric text from Excel cell C# | CellValueFormatStrategy.None example | Retrieve unformatted numeric value Aspose.Cells
// Developer Intent: The developer needs the exact numeric representation stored in an Excel cell, bypassing any applied number formats, using Aspose.Cells for .NET.
// Use Cases: Export the precise numeric text to a downstream system that requires the original string format. | Perform value comparisons before any custom number formatting is applied. | Log cell contents for audit trails or debugging without format-induced alterations.
// AI Prompts: Show how to extract raw numeric strings from a range of cells using GetStringValue(CellValueFormatStrategy.None) in Aspose.Cells for .NET. | Create a reusable C# method that returns unformatted string values for any given cell range. | Explain the differences between CellValueFormatStrategy.None and other strategies, and suggest when each should be used.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, writes the number 12345.6789 to cell A1, then uses GetStringValue(CellValueFormatStrategy.None) to obtain the exact numeric text without any formatting, displays it, and saves the workbook.
    public class GetRawNumericStringDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Put a numeric value into cell A1
                cells["A1"].PutValue(12345.6789);

                // Retrieve the raw numeric string without any formatting
                // CellValueFormatStrategy.None means no formatting is applied
                string rawNumericString = cells["A1"].GetStringValue(CellValueFormatStrategy.None);

                // Display the raw string value
                Console.WriteLine("Raw numeric string (no formatting): " + rawNumericString);

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                string outputPath = "GetRawNumericStringDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            GetRawNumericStringDemo.Run();
        }
    }
}
