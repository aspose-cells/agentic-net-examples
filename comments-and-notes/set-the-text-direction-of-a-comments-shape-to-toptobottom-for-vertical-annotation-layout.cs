// Title: Set a comment shape’s text direction to vertical (Top‑to‑Bottom) with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set comment shape text direction to vertical | how to display Excel comment text top to bottom using Aspose.Cells | use RightToLeft TextDirection for vertical comment layout in Aspose.Cells .NET | resize comment shape after applying vertical text direction in Aspose.Cells | vertical annotation layout for comments in Aspose.Cells workbook
// Tags: comment shape text direction Aspose.Cells | vertical comment layout Aspose.Cells | RightToLeft text orientation comment Aspose.Cells | resize comment shape for vertical text Aspose.Cells | Aspose.Cells comment orientation workaround

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SetCommentTextDirection
{
    // The example creates a workbook, adds a comment to cell B2, sets the comment note, accesses the CommentShape, assigns TextDirection = RightToLeft to simulate a top‑to‑bottom vertical layout, optionally resizes the shape for better fit, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output workbook
            const string outputPath = "CommentTextDirectionDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Put some sample data into a cell (optional, just for context)
                worksheet.Cells["B2"].PutValue("Sample Data");

                // 3. Add a comment to cell B2
                //    The Add method returns the index of the newly created comment.
                int commentIndex = worksheet.Comments.Add("B2");
                Comment comment = worksheet.Comments[commentIndex];

                // 4. Set the comment text (Note)
                comment.Note = "This comment is displayed vertically (Top‑to‑Bottom).";

                // 5. Access the shape that represents the comment.
                //    CommentShape is located in the Aspose.Cells.Drawing namespace.
                CommentShape commentShape = comment.CommentShape;

                // 6. Set the text direction of the comment shape.
                //    The TextDirectionType enum does not contain a TopToBottom value;
                //    using RightToLeft provides a vertical-like layout for comments.
                commentShape.TextDirection = TextDirectionType.RightToLeft;

                // 7. (Optional) Adjust the size of the comment shape so the vertical text fits nicely.
                commentShape.Width = 150;   // width in points
                commentShape.Height = 200;  // height in points

                // 8. Save the workbook to the specified file.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
                Console.WriteLine("The comment's text direction is set to RightToLeft (vertical layout).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
