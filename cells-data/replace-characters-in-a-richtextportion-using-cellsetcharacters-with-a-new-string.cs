// Title: Replace part of a rich‑text cell with SetCharacters in Aspose.Cells (C#)
// Description: Creates a workbook, writes "Hello World!" to A1, formats the word "World" as bold red, then substitutes that substring with "Aspose" while re‑applying the original FontSetting using Cell.Characters (SetCharacters). The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | SetCharacters | replace rich text | preserve formatting | C# | .NET | Excel cell characters | FontSetting | update substring | rich‑text API
// Common Searches: Aspose.Cells replace part of cell text | SetCharacters keep original font style | C# replace rich‑text substring in Excel | how to preserve formatting when changing cell text Aspose | update cell characters without losing style
// Developer Intent: Replace a specific substring in a cell’s rich‑text value and apply the original font attributes to the new text using Aspose.Cells.
// Use Cases: Change a highlighted keyword in a report header while retaining its bold red style. | Update product names in a spreadsheet without losing cell‑level rich‑text formatting. | Programmatically modify a phrase in a cell and keep its size, font, and color attributes.
// AI Prompts: Provide C# code that uses Aspose.Cells SetCharacters to replace a substring in a cell and copy the original FontSetting to the new characters. | Show an example that changes "World" to a variable string in cell A1, preserving bold and color formatting with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRichTextReplaceDemo
{
    // Creates a workbook, writes "Hello World!" to A1, formats the word "World" as bold red, then substitutes that substring with "Aspose" while re‑applying the original FontSetting using Cell.Characters (SetCharacters). The workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];

                // Original rich‑text value
                cell.Value = "Hello World!";

                // Apply formatting to the part we will replace ("World")
                FontSetting originalPart = cell.Characters(6, 5); // "World"
                originalPart.Font.IsBold = true;
                originalPart.Font.Color = Color.Red;

                // New text to insert
                string newText = "Aspose";

                // Build the new full string
                string original = cell.StringValue; // "Hello World!"
                string updated = original.Substring(0, 6) + newText + original.Substring(11);
                cell.Value = updated; // "Hello Aspose!"

                // Apply the original formatting to the replaced portion
                FontSetting newPart = cell.Characters(6, newText.Length);
                newPart.Font.IsBold = originalPart.Font.IsBold;
                newPart.Font.Color = originalPart.Font.Color;
                newPart.Font.Size = originalPart.Font.Size;
                newPart.Font.Name = originalPart.Font.Name;

                // Save the workbook
                workbook.Save("RichTextReplaceResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
