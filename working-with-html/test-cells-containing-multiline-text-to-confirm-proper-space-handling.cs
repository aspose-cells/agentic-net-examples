// Title: Verify preservation of leading and trailing spaces in multi‑line cell text using Aspose.Cells for .NET
// Common Searches: how to keep leading spaces in a multi line cell with Aspose.Cells C# | Aspose.Cells verify that spaces are not trimmed when reading wrapped cell text | C# read cell string value with original spacing after saving workbook using Aspose.Cells | preserve whitespace in Excel cell using Aspose.Cells .NET API | compare original multi line string with cell value in Aspose.Cells workbook
// Tags: preserve leading and trailing spaces in Excel cell using Aspose.Cells | set text wrap style for multi line cell Aspose.Cells | read cell string value with original whitespace Aspose.Cells | validate space handling after workbook save Aspose.Cells | compare multi line string length with cell value C#

using System;
using Aspose.Cells;

namespace AsposeCellsMultiLineTest
{
    // The example creates a workbook, writes a multi‑line string that includes leading, trailing, and inter‑line spaces into cell A1, enables text wrapping, saves the file, then reads the cell's string value back and compares its length to the original to confirm that whitespace is preserved.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (using the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define multi‑line text with intentional spaces before, after and between lines
            string multiLineText = "First line  \n  Second line   \nThird line";

            // Set the multi‑line text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(multiLineText);

            // Ensure the cell style wraps text so line breaks are visible in Excel
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Save the workbook (using the provided saving rule)
            workbook.Save("MultiLineTest.xlsx");

            // Retrieve the value back from the cell to verify space handling
            string retrievedText = cell.StringValue;

            // Output the original and retrieved texts for comparison
            Console.WriteLine("Original Text:");
            Console.WriteLine(multiLineText);
            Console.WriteLine("\nRetrieved Text:");
            Console.WriteLine(retrievedText);

            // Simple verification: check that the retrieved text length matches the original
            if (retrievedText.Length == multiLineText.Length)
            {
                Console.WriteLine("\nSpace handling verification passed: lengths match.");
            }
            else
            {
                Console.WriteLine("\nSpace handling verification failed: lengths differ.");
            }
        }
    }
}
