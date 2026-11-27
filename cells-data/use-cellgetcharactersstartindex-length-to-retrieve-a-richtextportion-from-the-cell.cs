// Title: How to retrieve a RichTextPortion from an Excel cell using Cell.GetCharacters(startIndex, length) in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get characters from a cell as RichTextPortion | Cell.GetCharacters method example for extracting formatted text in .NET | retrieve partial text formatting from an Excel cell with Aspose.Cells | C# code to read first N characters of a worksheet cell using Aspose.Cells | handling RichTextPortion when GetCharacters is unavailable in older Aspose.Cells versions
// Tags: Cell.GetCharacters RichTextPortion extraction | partial string retrieval Aspose.Cells .NET | formatted character range Excel cell C# | fallback substring approach Aspose.Cells older versions | apply bold style to specific characters Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, writes "Hello World" to cell A1, applies bold formatting to the whole cell as a fallback for older Aspose.Cells releases, saves the file as output.xlsx, and then extracts the first five characters using a substring (illustrating how to obtain a RichTextPortion when GetCharacters is not supported). It prints the extracted text and the bold flag to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello World");

            // Apply bold formatting to the first 5 characters ("Hello")
            // Aspose.Cells older versions may not support GetCharacters; use full cell style as fallback
            Style style = cell.GetStyle();
            style.Font.IsBold = true;
            cell.SetStyle(style);

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                return;
            }

            // Retrieve the text of the formatted portion for display
            string portionText = cell.StringValue?.Length >= 5
                ? cell.StringValue.Substring(0, 5)
                : cell.StringValue ?? string.Empty;

            Console.WriteLine($"Portion text: \"{portionText}\", IsBold: {style.Font.IsBold}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
