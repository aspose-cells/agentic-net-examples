// Title: Replace a word in a cell's rich text by updating the string value and reapplying formatting with Aspose.Cells for .NET
// Common Searches: how to replace a substring in a rich text cell using Aspose.Cells C# | Aspose.Cells replace word in cell while preserving existing formatting | C# example of updating cell characters when Cell.SetCharacters overload is missing | modify specific characters in an Excel cell with Aspose.Cells .NET | replace "World" with "Aspose" in cell A1 using Aspose.Cells C#
// Tags: rich text substring replacement Aspose.Cells | preserve existing cell formatting Aspose.Cells | manual cell string update Aspose.Cells | Cell.Characters API formatting Aspose.Cells | Excel .xlsx word replace Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRichTextReplace
{
    // The example creates a workbook, writes "Hello World!" to cell A1, formats the word "Hello" in bold blue, then replaces the substring "World" with "Aspose" by manipulating the cell's string value, applies italic green formatting to the new word, and saves the file as RichTextReplaceDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Target cell
                Cell cell = worksheet.Cells["A1"];

                // Set initial rich text value
                cell.PutValue("Hello World!");

                // Apply formatting to the first word "Hello"
                FontSetting helloFont = cell.Characters(0, 5);
                helloFont.Font.IsBold = true;
                helloFont.Font.Color = Color.Blue;
                helloFont.Font.Size = 12;

                // Replace the word "World" (starting at index 6, length 5) with "Aspose"
                // Since Cell.SetCharacters overload is unavailable, perform manual replacement
                string originalText = cell.StringValue;
                string newText = originalText.Remove(6, 5).Insert(6, "Aspose");
                cell.PutValue(newText);

                // Optionally, format the newly inserted word
                FontSetting asposeFont = cell.Characters(6, 6); // "Aspose" length is 6
                asposeFont.Font.IsItalic = true;
                asposeFont.Font.Color = Color.Green;
                asposeFont.Font.Size = 12;

                // Output the final cell text to console
                Console.WriteLine("Final cell text: " + cell.StringValue);

                // Save the workbook to verify the result
                string outputPath = "RichTextReplaceDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
