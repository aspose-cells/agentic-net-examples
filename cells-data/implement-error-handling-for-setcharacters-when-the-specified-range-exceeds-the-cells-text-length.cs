// Title: Validate character range and handle errors when using SetCharacters in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to check character range before calling SetCharacters | Aspose.Cells SetCharacters throws ArgumentOutOfRangeException when range exceeds cell text | example of error handling for SetCharacters method in Aspose.Cells .NET | prevent out of range error while replacing characters in an Excel cell with Aspose.Cells | validate start index and length for cell text in Aspose.Cells C#
// Tags: SetCharacters range validation Aspose.Cells | ArgumentOutOfRangeException handling Aspose.Cells C# | replace characters in Excel cell Aspose.Cells .NET | cell text length check Aspose.Cells | error handling for Excel cell modification Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, writes "Hello Aspose" to cell A1, verifies that the zero‑based start index and length do not exceed the existing cell text, throws an ArgumentOutOfRangeException for invalid ranges, replaces the specified characters with "World" when valid, saves the file, and logs any exceptions.
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

            // Put initial text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello Aspose");

            // Define the character range to replace
            int startIndex = 6;   // zero‑based start position
            int length = 10;      // number of characters to replace
            string newText = "World";

            // Retrieve the current text of the cell
            string currentText = cell.StringValue ?? string.Empty;

            // Validate that the specified range does not exceed the text length
            if (startIndex < 0 || length < 0 || startIndex + length > currentText.Length)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(startIndex),
                    "The specified character range exceeds the cell's text length.");
            }

            // Replace the specified character range with new text
            string updatedText = currentText.Remove(startIndex, length).Insert(startIndex, newText);
            cell.PutValue(updatedText);

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
