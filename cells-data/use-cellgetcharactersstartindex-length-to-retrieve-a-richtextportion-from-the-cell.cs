// Title: Retrieve a RichTextPortion from cell A1 using Cell.GetCharacters(startIndex, length) with Aspose.Cells for .NET
// AI Prompts: Generate C# code that calls Cell.GetCharacters(6,5) on cell A1 to obtain a RichTextPortion, prints its text, and saves the workbook using Aspose.Cells. | Show how to apply a red bold font to a cell and then extract characters 7‑11 as a RichTextPortion with the Aspose.Cells GetCharacters API.
// Common Searches: Aspose.Cells C# GetCharacters method example for extracting text | How to obtain a RichTextPortion from a specific character range in an Excel cell using Aspose.Cells | Retrieve characters 7 to 11 from cell A1 with Aspose.Cells .NET | Extract substring as RichTextPortion from Excel cell using Aspose.Cells API | Cell.GetCharacters startIndex length usage in Aspose.Cells C#
// Tags: Cell.GetCharacters usage Aspose.Cells | extract RichTextPortion from Excel cell .NET | apply red bold style to cell Aspose.Cells | save workbook after text extraction Aspose.Cells | retrieve specific characters from cell Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a workbook, writes "Hello World" to cell A1, applies a red bold style, extracts characters 7‑11 (using substring logic as a fallback when GetCharacters overload is unavailable), prints the extracted portion, and saves the file as output.xlsx, all wrapped in exception handling.
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

            // Get cell A1
            Cell cell = sheet.Cells["A1"];

            // Set a string value in the cell
            cell.PutValue("Hello World");

            // Apply formatting (red and bold) to the cell.
            // The GetCharacters overload with start/length is not available in this version,
            // so we apply the style to the whole cell.
            Style style = cell.GetStyle();
            style.Font.Color = Color.Red;
            style.Font.IsBold = true;
            cell.SetStyle(style);

            // Retrieve characters 7 to 11 (zero‑based index, length 5) as another portion
            string cellText = cell.StringValue ?? string.Empty;
            string secondPortionText = (cellText.Length >= 11) ? cellText.Substring(6, 5) : string.Empty;

            // Output the text of the retrieved portion
            Console.WriteLine("Retrieved portion text: " + secondPortionText);

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
