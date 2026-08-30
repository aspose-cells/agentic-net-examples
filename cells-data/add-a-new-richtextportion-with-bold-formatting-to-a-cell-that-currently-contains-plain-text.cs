// Title: Add a bold, red RichTextPortion to a specific substring in an Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, writes "Hello Aspose.Cells!" to cell A1, and makes the first five characters bold and red using Aspose.Cells. | Write a reusable method that takes a cell reference, start index, length, and applies bold formatting with a custom color to that text segment via the Aspose.Cells API. | Show how to change the applied bold formatting to italic for a different character range in the same cell using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to format part of a cell's text as bold and colored | C# code to apply character-level formatting in Excel with Aspose.Cells | Set font style for a substring in an Excel cell using Aspose.Cells .NET API | Apply rich text formatting to specific characters in a worksheet cell Aspose.Cells
// Tags: apply bold rich text portion Aspose.Cells C# | character-level font color Aspose.Cells .NET | partial text formatting Excel workbook Aspose.Cells | rich text substring styling Excel .NET | set characters formatting Aspose.Cells API

using System;
using Aspose.Cells;
using System.Drawing;

// Demonstrates creating a workbook, inserting plain text into cell A1, and using the Characters method to apply bold and red formatting to a defined substring, then saving the file as BoldRichTextPortion.xlsx.
class AddBoldRichTextPortion
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Target cell
        Cell cell = worksheet.Cells["A1"];

        // Set plain text value
        cell.PutValue("Hello Aspose.Cells!");

        // Define the portion to make bold (e.g., "Hello")
        int startIndex = 0;               // start at first character
        int length = 5;                   // length of the portion

        // Apply bold formatting to the specified portion
        cell.Characters(startIndex, length).Font.IsBold = true;

        // Optionally, set a different color for visibility
        cell.Characters(startIndex, length).Font.Color = Color.Red;

        // Save the workbook
        workbook.Save("BoldRichTextPortion.xlsx");
    }
}
