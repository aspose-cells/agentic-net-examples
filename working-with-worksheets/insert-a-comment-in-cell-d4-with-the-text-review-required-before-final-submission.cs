// Title: Insert a comment into cell D4 of an Excel worksheet and save with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds the note "Review required before final submission" to the worksheet cell located at row 4, column D, removes any existing comment first, and saves the workbook as Output.xlsx using Aspose.Cells. | Show how to detect and delete a comment in a particular cell before creating a new comment with Aspose.Cells in a .NET application. | Provide a step‑by‑step example of adding a comment to a worksheet cell, handling existing comments, and persisting the file with Aspose.Cells.
// Common Searches: Aspose.Cells C# add comment to cell D4 and overwrite existing comment | how to remove a comment from a specific Excel cell using Aspose.Cells .NET | C# example for inserting a note into an Excel worksheet cell with Aspose.Cells | save Excel file with comments using Aspose.Cells for .NET | checking for existing comments before adding new comment Aspose.Cells C#
// Tags: worksheet cell comment insertion Aspose.Cells | existing comment removal Aspose.Cells C# | comment management in Aspose.Cells .NET | saving workbook with comments Aspose.Cells | note insertion into specific worksheet cell

using Aspose.Cells;

// Creates a new Workbook, ensures the target cell (row 4, column D) has no prior comment, adds a comment with the text "Review required before final submission", and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the target cell (D4 => row 3, column 3, zero‑based indexing)
        int row = 3;
        int column = 3;

        // Ensure no existing comment occupies the cell
        if (sheet.Comments[row, column] != null)
            sheet.Comments.RemoveAt(row, column);

        // Add a new comment to the cell
        int commentIndex = sheet.Comments.Add(row, column);
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "Review required before final submission";

        // Save the workbook (lifecycle save)
        workbook.Save("Output.xlsx");
    }
}
