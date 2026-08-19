// Title: Aspose.Cells C# – Validate Ranges and Handle Errors for Cell.SetCharacters
// Description: Demonstrates how to check each (start, length) pair against the actual text in a worksheet cell, apply FontSetting formatting, and safely invoke Cell.SetCharacters with try‑catch blocks. Includes error handling for out‑of‑range ranges and for saving the workbook.
// Keywords: Aspose.Cells SetCharacters validation | Cell.Characters range check C# | partial text formatting Excel | ArgumentOutOfRangeException handling | workbook.Save error handling
// Common Searches: C# Aspose.Cells SetCharacters range error | prevent SetCharacters exception Aspose | validate start length for cell characters | safe partial formatting of Excel cell
// Developer Intent: Add pre‑validation of character ranges and surround SetCharacters and workbook.Save with exception handling to avoid runtime failures.
// Use Cases: Bold the first five characters of a cell while confirming the range fits the string length. | Iterate over multiple (start, length) tuples, apply distinct fonts to each segment, and catch any formatting errors. | Log or display meaningful messages when saving the Excel file fails after applying character formatting.
// AI Prompts: Generate C# code that verifies each (start, length) tuple against a cell's text length before calling cell.SetCharacters in Aspose.Cells, throwing an ArgumentOutOfRangeException for invalid entries. | Create a robust Aspose.Cells example that formats several character ranges in a cell, includes range validation, and demonstrates try‑catch handling for both SetCharacters and workbook.Save.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to check each (start, length) pair against the actual text in a worksheet cell, apply FontSetting formatting, and safely invoke Cell.SetCharacters with try‑catch blocks. Includes error handling for out‑of‑range ranges and for saving the workbook.
    public class SetCharactersWithValidationDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set the cell value that will be formatted
            string cellText = "Hello World!";
            cell.PutValue(cellText);

            // Define the character ranges we want to format
            // Example: first 5 characters bold red, next 6 characters italic blue
            var ranges = new (int start, int length)[]
            {
                (0, 5),   // "Hello"
                (6, 6)    // "World!"
            };

            // Prepare the FontSetting array after validating each range
            FontSetting[] fontSettings = new FontSetting[ranges.Length];

            for (int i = 0; i < ranges.Length; i++)
            {
                int start = ranges[i].start;
                int length = ranges[i].length;

                // Validate the range against the actual cell text length
                if (start < 0 || length < 0 || start + length > cellText.Length)
                {
                    throw new ArgumentOutOfRangeException(
                        $"The range (start={start}, length={length}) exceeds the cell text length ({cellText.Length}).");
                }

                // Obtain the FontSetting for the valid range
                fontSettings[i] = cell.Characters(start, length);

                // Apply sample formatting based on the range index
                if (i == 0)
                {
                    fontSettings[i].Font.IsBold = true;
                    fontSettings[i].Font.Color = Color.Red;
                }
                else
                {
                    fontSettings[i].Font.IsItalic = true;
                    fontSettings[i].Font.Color = Color.Blue;
                }
            }

            // Apply the formatted character settings using SetCharacters with error handling
            try
            {
                cell.SetCharacters(fontSettings);
                Console.WriteLine("SetCharacters executed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing SetCharacters: {ex.Message}");
            }

            // Save the workbook
            try
            {
                workbook.Save("SetCharactersWithValidationDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SetCharactersWithValidationDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
