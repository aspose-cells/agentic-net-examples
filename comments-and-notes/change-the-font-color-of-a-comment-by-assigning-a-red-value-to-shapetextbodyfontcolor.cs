// Title: How to set the font color of an Excel cell comment to red using Aspose.Cells for .NET (C#)
// AI Prompts: Change the comment text color to red by accessing the CommentShape.TextBody[0].Font.Color property in C# with Aspose.Cells. | Apply a custom font color to an Excel comment by modifying the first FontSetting in the comment's TextBody using Aspose.Cells .NET. | Programmatically set the font color of a worksheet comment and save the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set comment font color to red | How to change the text color of an Excel comment using Aspose.Cells | Modify comment shape text body font color programmatically in .NET | Change Excel cell comment text color with Aspose.Cells API
// Tags: Aspose.Cells comment shape font color | C# set comment text color Aspose.Cells | CommentShape TextBody Font.Color example | Excel comment font color Aspose.Cells .NET | programmatic comment styling Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentFontColor
{
    // Demonstrates creating a workbook, adding a comment to cell A1, retrieving its CommentShape, setting the comment text font color to red via TextBody[0].Font.Color, and saving the file as CommentFontColorDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Get the shape associated with the comment
            CommentShape commentShape = comment.CommentShape;

            // Change the font color of the comment text via the shape's TextBody
            // TextBody is a collection of FontSetting objects; we modify the first (and only) entry
            commentShape.TextBody[0].Font.Color = Color.Red;

            // Save the workbook (save rule)
            workbook.Save("CommentFontColorDemo.xlsx");
        }
    }
}
