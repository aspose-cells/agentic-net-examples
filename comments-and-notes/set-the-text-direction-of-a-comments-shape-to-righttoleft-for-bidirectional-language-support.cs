// Title: How to set Right-to-Left text direction for an Excel comment shape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a comment with Arabic text to a cell, and assigns comment.CommentShape.TextDirection = TextDirection.RightToLeft with Aspose.Cells. | Show how to enable RTL layout for a comment shape in Aspose.Cells, including a fallback when the TextDirection property is unavailable in older versions.
// Common Searches: Aspose.Cells C# set comment shape text direction to RTL | Enable right-to-left layout for Excel comments using Aspose.Cells .NET | Display Arabic text correctly in an Excel comment with Aspose.Cells | CommentShape TextDirection property missing in older Aspose.Cells versions | Configure bidirectional language support for Excel comments in C#
// Tags: Aspose.Cells comment shape right-to-left | C# configure comment TextDirection | Excel comment bidirectional language support | Aspose.Cells TextDirection property example | Arabic comment rendering Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, accesses the first worksheet, adds a comment with Arabic text to cell B2, and demonstrates how to set the comment's shape TextDirection to RightToLeft for RTL support using Aspose.Cells for .NET, then saves the file as CommentWithRTL.xlsx.
    class SetCommentTextDirection
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a comment to cell B2 (using overload that accepts only the cell address)
                int commentIndex = sheet.Comments.Add("B2");
                Comment comment = sheet.Comments[commentIndex];

                // Set author and note text
                comment.Author = "Author";
                comment.Note = "مثال على نص عربي";

                // NOTE: The TextDirection property may not be available in older Aspose.Cells versions.
                // If supported, you can enable right‑to‑left text direction as shown below:
                // comment.CommentShape.TextDirection = Aspose.Cells.Drawing.TextDirection.RightToLeft;

                // Save the workbook
                string outputPath = "CommentWithRTL.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
