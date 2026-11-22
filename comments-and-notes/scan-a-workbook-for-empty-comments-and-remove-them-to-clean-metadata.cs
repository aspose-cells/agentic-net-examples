// Title: How to remove empty cell comments from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove comments with no text | Delete blank cell notes from an Excel file using Aspose.Cells API | Iterate comment collection and clean empty entries in .NET | How to clean up empty comments in an Excel workbook with Aspose.Cells | Remove whitespace‑only notes from worksheets using Aspose.Cells C#
// Tags: remove empty comments Aspose.Cells | Aspose.Cells comment collection cleanup | delete blank cell notes .NET | iterate worksheet comments C# | clean Excel metadata Aspose.Cells | filter null or whitespace notes Excel

using System;
using Aspose.Cells;

// The example loads an Excel workbook, loops through each worksheet's CommentCollection, removes any comment whose Note property is null, empty, or whitespace, and saves the cleaned workbook to a new file.
class RemoveEmptyComments
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path to save the cleaned workbook
        string outputPath = "output.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of comments for the current worksheet
            CommentCollection comments = sheet.Comments;

            // Iterate backwards so that removal does not affect the index order
            for (int i = comments.Count - 1; i >= 0; i--)
            {
                Comment comment = comments[i];

                // If the comment text is null, empty, or whitespace, remove it
                if (string.IsNullOrWhiteSpace(comment.Note))
                {
                    comments.RemoveAt(i);
                }
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save(outputPath);
    }
}
