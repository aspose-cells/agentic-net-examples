// Title: Add a Bold RichTextPortion to an Existing Cell with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes plain text to cell A1, inserts additional text, and applies bold formatting to the new RichTextPortion using the Characters method before saving the file.
// Keywords: Aspose.Cells | C# rich text | Bold RichTextPortion | InsertText method | Characters formatting | Excel cell bold text | Aspose.Cells .NET | RichTextPortion formatting | Add bold text to cell
// Common Searches: Aspose.Cells add bold text to part of a cell | C# insert RichTextPortion and make it bold | How to format a substring in an Excel cell using Aspose.Cells | Apply bold style to specific characters with Aspose.Cells for .NET | Insert text and set font weight in a worksheet cell
// Developer Intent: Insert a new RichTextPortion into a cell that already contains plain text and set that portion to bold.
// Use Cases: Create a report header where the label is regular and the key term is bold within the same cell. | Generate an invoice line where the description stays normal but the amount appears in bold after insertion. | Highlight keywords in a paragraph by inserting them as bold RichTextPortions in a single Excel cell.
// AI Prompts: Show C# code that adds a bold RichTextPortion to an existing cell using Aspose.Cells. | Provide an example of inserting multiple RichTextPortions with different styles (bold, italic, color) into one cell. | Explain how to change the font weight of a specific character range after using InsertText with Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, writes plain text to cell A1, inserts additional text, and applies bold formatting to the new RichTextPortion using the Characters method before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the target cell and set initial plain text
        Cell cell = worksheet.Cells["A1"];
        string plainText = "Hello ";
        cell.PutValue(plainText);

        // Text that will be added as a new rich‑text portion
        string boldPortion = "World";

        // Insert the new text at the end of the existing content
        cell.InsertText(plainText.Length, boldPortion);

        // Apply bold formatting to the inserted portion
        cell.Characters(plainText.Length, boldPortion.Length).Font.IsBold = true;

        // Save the workbook
        workbook.Save("RichTextBold.xlsx");
    }
}
