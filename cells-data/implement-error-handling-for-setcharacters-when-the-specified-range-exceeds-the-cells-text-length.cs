// Title: Add error handling for Cell.Characters SetCharacters when the specified character range exceeds the cell text length in Aspose.Cells for .NET
// AI Prompts: Check that startIndex and length do not exceed cell.StringValue.Length before calling cell.Characters, and output a clear error message if they do. | Wrap the cell.SetCharacters call in a try‑catch block to log any exception that occurs during rich‑text formatting. | If the requested range is larger than the text, truncate the length to the maximum possible value and then apply the formatting.
// Common Searches: Aspose.Cells .NET prevent SetCharacters out of range error | C# validate character range before using Cell.Characters in Aspose.Cells | example of try catch around SetCharacters for rich text formatting | how to handle SetCharacters exception when start index is beyond cell text length
// Tags: Aspose.Cells SetCharacters range validation | Cell.Characters length check .NET | rich text formatting error handling Aspose.Cells | SetCharacters exception handling C# | Aspose.Cells workbook save after formatting

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to validate the start index and length against the cell's text, apply rich‑text formatting with SetCharacters inside a try‑catch block, and handle cases where the character range exceeds the cell content before saving the workbook.
    public class SetCharactersErrorHandlingDemo
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access cell A1 and set a short text value
            Cell cell = worksheet.Cells["A1"];
            cell.Value = "Hello";

            // Define the character range we want to format
            int startIndex = 0;
            int length = 10; // Intentionally exceeds the text length (5)

            // Retrieve the current cell text
            string cellText = cell.StringValue ?? string.Empty;

            // Prepare an array to hold FontSetting objects
            FontSetting[] fontSettings = new FontSetting[1];

            // Validate the range before calling SetCharacters
            if (startIndex < 0 || length < 0 || startIndex + length > cellText.Length)
            {
                Console.WriteLine("Error: Specified character range exceeds the cell's text length.");
                Console.WriteLine($"Cell text length: {cellText.Length}, Requested range: start={startIndex}, length={length}");
            }
            else
            {
                // Create the FontSetting for the valid range
                fontSettings[0] = cell.Characters(startIndex, length);
                fontSettings[0].Font.IsBold = true;
                fontSettings[0].Font.Color = Color.Blue;

                // Apply the rich text formatting using SetCharacters with error handling
                try
                {
                    cell.SetCharacters(fontSettings);
                    Console.WriteLine("SetCharacters executed successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception during SetCharacters: {ex.Message}");
                }
            }

            // Save the workbook
            workbook.Save("SetCharactersErrorHandlingDemo.xlsx");
        }
    }
}
