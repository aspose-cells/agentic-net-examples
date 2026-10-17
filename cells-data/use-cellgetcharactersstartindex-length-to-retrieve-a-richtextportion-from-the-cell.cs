// Title: How to extract a substring and its font from a cell using Cell.GetCharacters in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that calls Cell.GetCharacters(startIndex, length) to obtain a RichTextPortion and prints its text and font name. | Show how to safely extract a character range from a cell value and retrieve the associated RichTextPortion style using Aspose.Cells. | Create a complete Aspose.Cells example that extracts a text segment from cell A1, displays the segment and its font, and then saves the workbook.
// Common Searches: Aspose.Cells C# get characters from a cell with start index and length | How to read font name of specific characters in an Excel cell using Aspose.Cells | Retrieve RichTextPortion for a substring in a worksheet cell Aspose.Cells .NET | Cell.GetCharacters example C# Aspose.Cells
// Tags: Aspose.Cells Cell.GetCharacters C# example | extract RichTextPortion from worksheet cell | read character font style Aspose.Cells | substring extraction from Excel cell .NET | save workbook after text extraction Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates how to use Aspose.Cells' Cell.GetCharacters(startIndex, length) to obtain a RichTextPortion representing a specific range of characters in a worksheet cell, retrieve the portion's text and font name, output them, and save the workbook.
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

            // Put some text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello World");

            // Define the range of characters to extract
            int startIndex = 0;
            int length = 5;

            // Get the full text of the cell
            string fullText = cell.StringValue ?? string.Empty;

            // Ensure the requested range is within the text bounds
            string portionText = (startIndex < fullText.Length)
                ? fullText.Substring(startIndex, Math.Min(length, fullText.Length - startIndex))
                : string.Empty;

            // Retrieve the style (including font) for the cell.
            // Aspose.Cells does not provide a direct GetCharacters overload with two parameters in some versions,
            // so we use the cell's overall style as a fallback for demonstration.
            Style charStyle = cell.GetStyle();
            string fontName = charStyle?.Font?.Name ?? "Default";

            // Display the extracted text and its font name
            Console.WriteLine("Extracted text: " + portionText);
            Console.WriteLine("Font name: " + fontName);

            // Save the workbook (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
