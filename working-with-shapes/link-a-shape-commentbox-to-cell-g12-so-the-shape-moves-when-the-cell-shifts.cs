// Title: How to link a CommentBox shape to cell G12 in Aspose.Cells for .NET so it follows row insertions
// Common Searches: Aspose.Cells C# link comment shape to a specific cell | move comment box automatically when inserting rows with Aspose.Cells | set LinkedCell property for CommentShape in a .NET workbook | bind Excel comment box to cell G12 using Aspose.Cells API | preserve comment position after row insertion in C# Aspose.Cells
// Tags: CommentShape LinkedCell Aspose.Cells .NET | bind Excel comment to cell programmatically | automatic comment reposition after row insert | Aspose.Cells dynamic comment positioning | C# set comment shape linked cell

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace CommentBoxLinkDemo
{
    // The example creates a workbook, adds a comment to cell G12, obtains its CommentShape, explicitly sets the shape's LinkedCell to "G12", inserts a row before the original row 12 (causing the linked comment box to shift automatically), and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "CommentBoxLinkedToG12.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Add a comment to cell G12 and obtain its CommentShape
                // ------------------------------------------------------------
                int commentIndex = sheet.Comments.Add("G12");
                Comment comment = sheet.Comments[commentIndex];
                comment.Note = "This is a linked comment box.";

                // The shape that represents the comment box
                CommentShape commentBox = comment.CommentShape;

                if (commentBox == null)
                {
                    Console.WriteLine("Failed to retrieve the CommentShape.");
                    return;
                }

                // ------------------------------------------------------------
                // 2. Link the comment box to cell G12 (already linked, but set explicitly)
                // ------------------------------------------------------------
                commentBox.LinkedCell = "G12";
                Console.WriteLine($"CommentBox linked to cell: {commentBox.LinkedCell}");

                // ------------------------------------------------------------
                // 3. Insert a row above row 12; the linked comment box moves automatically
                // ------------------------------------------------------------
                sheet.Cells.InsertRows(11, 1); // Inserts a row before the original row 12 (index 11)

                Console.WriteLine($"After inserting a row, CommentBox linked to cell: {commentBox.LinkedCell}");

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
