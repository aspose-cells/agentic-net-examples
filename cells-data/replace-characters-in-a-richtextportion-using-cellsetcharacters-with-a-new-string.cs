// Title: Replace a substring in a cell’s rich‑text and apply separate formatting with Cell.SetCharacters in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that finds a specific word inside an Excel cell, replaces it with a new word, and then uses FontSetting objects with Cell.SetCharacters to apply bold blue formatting to the first part and italic green formatting to the replacement. | Show how to update the text of a cell and then assign different font styles to distinct character ranges using Aspose.Cells’ Cell.Characters and Cell.SetCharacters methods.
// Common Searches: Aspose.Cells C# replace word in cell rich text and keep formatting | How to use Cell.SetCharacters to format parts of a string in Excel with Aspose | C# change substring in Excel cell and apply bold and italic colors using Aspose.Cells | Set different font styles for separate portions of a cell after updating its value in Aspose.Cells
// Tags: Cell.SetCharacters replace substring C# | Rich text portion formatting Aspose.Cells | FontSetting bold blue Excel cell | FontSetting italic green replacement text | Aspose.Cells update cell value .xlsx

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRichTextReplaceDemo
{
    // Demonstrates how to replace the word “World” with “Aspose” in cell A1, then apply bold blue formatting to “Hello ” and italic green formatting to “Aspose” using FontSetting objects and Cell.SetCharacters, and saves the workbook as RichTextReplaceOutput.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access cell A1 and set initial rich text value
            Cell cell = worksheet.Cells["A1"];
            cell.Value = "Hello World";

            // Replace the word "World" with "Aspose" by updating the cell value
            // (SetCharacters works on formatting, not on the actual text content)
            cell.Value = "Hello Aspose";

            // Prepare FontSetting objects for each portion of the new text
            // Portion 0: "Hello " (indices 0-5)
            FontSetting part1 = cell.Characters(0, 6);
            part1.Font.IsBold = true;
            part1.Font.Color = Color.Blue;

            // Portion 1: "Aspose" (indices 6-12)
            FontSetting part2 = cell.Characters(6, 6);
            part2.Font.IsItalic = true;
            part2.Font.Color = Color.Green;

            // Apply the formatting to the cell using SetCharacters
            FontSetting[] settings = new FontSetting[] { part1, part2 };
            cell.SetCharacters(settings);

            // Save the workbook
            workbook.Save("RichTextReplaceOutput.xlsx");
        }
    }
}
