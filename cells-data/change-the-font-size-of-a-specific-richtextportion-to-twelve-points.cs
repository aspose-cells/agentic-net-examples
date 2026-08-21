// Title: Set RichTextPortion font size to 12 pt in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes "Hello World" to A1, selects the characters "World" with Cell.Characters(6,5), changes the Font.Size to 12 points via the returned FontSetting, and saves the file as RichTextPortionFontSize.xlsx.
// Keywords: Aspose.Cells | C# | .NET | RichTextPortion | FontSetting | set font size | partial cell text | Excel font formatting | cell characters | 12 point font
// Common Searches: Aspose.Cells change font size of part of a cell | C# set RichTextPortion size in Excel | How to format specific characters in an Excel cell using Aspose | Set partial text font size with Aspose.Cells .NET | FontSetting.Size example Aspose.Cells
// Developer Intent: Modify the font size of the word "World" inside cell A1 to 12 pt using Aspose.Cells for .NET.
// Use Cases: Highlight keywords within a cell by enlarging their font. | Create styled headings where only selected words appear larger. | Generate reports that require emphasis on specific phrases without splitting cells.
// AI Prompts: Write C# code with Aspose.Cells to set the font size of characters 6‑10 in cell A1 to 12 pt. | Explain how the FontSetting object returned by Cell.Characters can be used to adjust size, color, and style of partial cell text. | Show an example that changes the font size of multiple non‑contiguous RichTextPortions in a worksheet using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, writes "Hello World" to A1, selects the characters "World" with Cell.Characters(6,5), changes the Font.Size to 12 points via the returned FontSetting, and saves the file as RichTextPortionFontSize.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put some text into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Hello World");

        // Obtain a RichTextPortion (FontSetting) for the word "World"
        // Start index is 6 (zero‑based) and length is 5 characters
        FontSetting richTextPortion = cell.Characters(6, 5);

        // Change the font size of this portion to twelve points
        richTextPortion.Font.Size = 12;

        // Save the workbook
        workbook.Save("RichTextPortionFontSize.xlsx");
    }
}
