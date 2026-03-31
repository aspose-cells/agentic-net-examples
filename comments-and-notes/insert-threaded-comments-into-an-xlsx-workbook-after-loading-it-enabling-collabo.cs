using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from disk
        Workbook workbook = new Workbook("InputFile.xlsx");

        // Access the first worksheet (you can change the index or use worksheet name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author to the workbook
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "Alice",                 // Author display name
            "alice@example.com",     // Author email or identifier
            "Provider1");            // Optional provider identifier
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Insert a threaded comment into cell B2
        worksheet.Comments.AddThreadedComment("B2", "Initial comment from Alice.", author);

        // Insert a reply (another threaded comment) to the same cell
        worksheet.Comments.AddThreadedComment("B2", "Follow‑up comment.", author);

        // Save the modified workbook
        workbook.Save("OutputFile.xlsx");
    }
}