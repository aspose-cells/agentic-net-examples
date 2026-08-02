// Title: Set Right-to-Left Text Direction for an Excel Comment Shape with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add an Arabic comment to cell A1, retrieve the associated CommentShape, set its TextDirection to RightToLeft, and save the file. This enables proper rendering of RTL languages such as Arabic or Hebrew in Excel comments.
// Keywords: Aspose.Cells | CommentShape | TextDirection | RightToLeft | RTL comment | Arabic Excel comment | Hebrew Excel comment | C# | .NET | bidirectional language support | Excel comment formatting
// Common Searches: Aspose.Cells set comment direction RTL | C# change Excel comment text direction | RightToLeft comment shape Aspose.Cells | How to display Arabic comment in Excel using Aspose | Set TextDirectionType for comment in .NET
// Developer Intent: Apply RightToLeft text direction to a comment's shape so RTL languages display correctly in Excel.
// Use Cases: Generate a report with Arabic or Hebrew comments that require RTL layout. | Batch‑process an existing workbook to convert all comment shapes to RightToLeft. | Create multilingual Excel templates where specific comments need bidirectional rendering.
// AI Prompts: Show C# code that sets CommentShape.TextDirection to RightToLeft with Aspose.Cells. | Provide a loop that updates TextDirection for every comment in a workbook to RightToLeft. | Explain the TextDirectionType options available for CommentShape in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentTextDirection
{
    // Demonstrates how to create a workbook, add an Arabic comment to cell A1, retrieve the associated CommentShape, set its TextDirection to RightToLeft, and save the file. This enables proper rendering of RTL languages such as Arabic or Hebrew in Excel comments.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "مثال على النص من اليمين إلى اليسار"; // Sample Arabic text

            // Access the shape associated with the comment
            CommentShape commentShape = comment.CommentShape;

            // Set the text direction of the comment's shape to RightToLeft
            commentShape.TextDirection = TextDirectionType.RightToLeft;

            // Save the workbook (lifecycle: save)
            workbook.Save("CommentRightToLeft.xlsx");
        }
    }
}
