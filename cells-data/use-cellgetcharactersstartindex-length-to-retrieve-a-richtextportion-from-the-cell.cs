// Title: Get and style a RichTextPortion with Cell.Characters in Aspose.Cells for .NET
// Description: C# sample that creates a workbook, writes "HelloWorld" to cell A1, extracts the characters at indexes 5‑9 via Cell.Characters, and applies bold blue formatting to that substring before saving the file.
// Keywords: Aspose.Cells | Cell.Characters | RichTextPortion | FontSetting | C# Excel formatting | partial cell text styling | substring formatting Aspose | Excel rich text API | apply bold color to part of cell | retrieve characters range
// Common Searches: Aspose.Cells retrieve rich text portion | Cell.Characters method example C# | format part of an Excel cell with Aspose | how to bold specific characters in a cell using Aspose.Cells | extract substring from cell and change font color
// Developer Intent: Extract a defined range of characters from a worksheet cell and apply custom font attributes to that segment.
// Use Cases: Emphasize a keyword inside a cell by making it bold and colored. | Display a monetary value with the currency symbol in a different style from the amount. | Separate date components (day, month, year) with distinct fonts within a single cell.
// AI Prompts: Show how to retrieve a RichTextPortion from cell B2 starting at index 3 with length 4 and set its font to italic and red using Aspose.Cells for .NET. | Generate C# code that scans column A and makes the word "Total" bold wherever it appears inside a cell using Cell.Characters. | Provide an example that underlines and colors the last three characters of a cell value green with Aspose.Cells.

using System;
using Aspose.Cells;
using System.Drawing;

// C# sample that creates a workbook, writes "HelloWorld" to cell A1, extracts the characters at indexes 5‑9 via Cell.Characters, and applies bold blue formatting to that substring before saving the file.
class RetrieveRichTextPortion
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a string value into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("HelloWorld");

        // Retrieve a rich‑text portion (FontSetting) for characters starting at index 5 with length 5 ("World")
        FontSetting richPortion = cell.Characters(5, 5);

        // Apply formatting to the retrieved portion
        richPortion.Font.IsBold = true;
        richPortion.Font.Color = Color.Blue;

        // Save the workbook
        workbook.Save("RichTextPortionDemo.xlsx");
    }
}
