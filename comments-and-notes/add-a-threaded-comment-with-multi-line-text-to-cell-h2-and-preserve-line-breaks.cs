// Title: Add a Multi‑Line Threaded Comment to Cell H2 with Line Breaks using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a threaded‑comment author, build a comment string that contains "\n" line breaks, attach the comment to cell H2, and save the file as ThreadedCommentMultiLine.xlsx.
// Keywords: Aspose.Cells C# | threaded comment Excel | multi line comment | preserve newline in comment | add comment to cell H2 | Excel API author comment | save workbook Aspose
// Common Searches: Aspose.Cells add threaded comment with line breaks | C# code to insert multi‑line comment in Excel | how to set author for Excel threaded comment using Aspose | preserve newline characters in Excel comment .NET | threaded comment example for cell H2
// Developer Intent: Insert a threaded comment that spans several lines into cell H2 while keeping the line breaks intact.
// Use Cases: Embed detailed review notes directly in generated reports. | Provide step‑by‑step instructions for a specific cell in a collaborative spreadsheet. | Mark cells with author‑attributed annotations that require formatted, multi‑line text.
// AI Prompts: Generate C# code with Aspose.Cells to add a threaded comment containing newline characters to cell H2 and assign a custom author. | Show how to create multiple threaded comments with different authors, each preserving its own line breaks. | Explain how to read, modify, or delete an existing multi‑line threaded comment in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, define a threaded‑comment author, build a comment string that contains "\n" line breaks, attach the comment to cell H2, and save the file as ThreadedCommentMultiLine.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "Demo Author",          // author name
            "demo@example.com",     // user id (email)
            "DemoProvider");        // provider id
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Define multi‑line comment text; "\n" preserves line breaks in Excel comments
        string multiLineText = "First line of comment.\nSecond line of comment.\nThird line.";

        // Add the threaded comment to cell H2 (row 1, column 7) using the cell name
        worksheet.Comments.AddThreadedComment("H2", multiLineText, author);

        // Save the workbook
        workbook.Save("ThreadedCommentMultiLine.xlsx");
    }
}
