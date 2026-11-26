// Title: Apply MS Mincho to Japanese characters and Arial to Latin characters in an Excel cell with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set different fonts for Japanese and English text in the same cell | How to use the Characters range to apply MS Mincho to Far East text in Excel with Aspose.Cells | C# Aspose.Cells example for mixed language font styling in a worksheet cell | Set Latin font to Arial and FarEast font to MS Mincho in an Aspose.Cells workbook | Apply rich‑text font changes to part of a cell using Aspose.Cells .NET
// Tags: mixed language font assignment Aspose.Cells | far east font MS Mincho Excel .NET | latin font Arial cell styling Aspose | rich text character range API Aspose.Cells | save workbook with multilingual fonts .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, writes "Hello こんにちは" to cell A1, applies Arial as the default Latin font to the whole cell, then uses the Characters range to set MS Mincho for the Japanese characters, and saves the file as MultilingualFontExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Multilingual text to place in cell A1
            const string text = "Hello こんにちは";

            // Set the text in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(text);

            // Create a style for Latin characters (Arial)
            Style latinStyle = workbook.CreateStyle();
            latinStyle.Font.Name = "Arial";

            // Apply the Latin style to the whole cell (font-related settings only)
            StyleFlag flag = new StyleFlag { Font = true };
            cell.SetStyle(latinStyle, flag);

            // Apply a different font (MS Mincho) to the Far East characters using rich text
            int startIndex = text.IndexOf('こ'); // first Japanese character
            if (startIndex >= 0)
            {
                int length = text.Length - startIndex;
                // Set the Far East font for the specific character range
                cell.Characters(startIndex, length).Font.Name = "MS Mincho";
            }

            // Define output file path
            string outputPath = "MultilingualFontExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
