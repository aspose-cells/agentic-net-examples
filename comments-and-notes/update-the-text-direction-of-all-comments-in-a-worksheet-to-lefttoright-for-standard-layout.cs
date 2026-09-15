// Title: Set all worksheet comment text directions to LeftToRight with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, loops through each comment in a worksheet, and assigns TextDirectionType.LeftToRight to comment.CommentShape.TextDirection. | Adapt an existing Aspose.Cells example to apply a left‑to‑right text direction to every comment shape in a workbook in a single pass. | Generate a C# snippet that changes the orientation of all comments in the first worksheet to left‑to‑right and saves the updated file.
// Common Searches: how to change comment text direction to left to right using Aspose.Cells C# | Aspose.Cells iterate through worksheet comments and set TextDirection | bulk update Excel comment orientation with Aspose.Cells .NET | C# set TextDirectionType.LeftToRight for all comments in a workbook | Aspose.Cells comment shape text direction left-to-right example
// Tags: Aspose.Cells set comment text direction | C# bulk update Excel comment orientation | Aspose.Cells comment shape TextDirectionType | modify worksheet comments left-to-right .NET | Excel comment text direction Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates over every comment in the first worksheet, sets each comment's TextDirection to LeftToRight via the CommentShape API, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the target worksheet (e.g., the first one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Update text direction for every comment in the worksheet
        foreach (Comment comment in worksheet.Comments)
        {
            // Set the comment's text direction to LeftToRight
            comment.CommentShape.TextDirection = TextDirectionType.LeftToRight;
        }

        // Save the changes
        workbook.Save("output.xlsx");
    }
}
