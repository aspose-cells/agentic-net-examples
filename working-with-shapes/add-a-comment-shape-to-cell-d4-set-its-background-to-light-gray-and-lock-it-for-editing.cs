using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell D4
            int commentIndex = worksheet.Comments.Add("D4");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "Sample comment";

            // Retrieve the shape attached to the comment
            Shape commentShape = comment.CommentShape;

            // Optional: set a background fill color (if supported by the current Aspose.Cells version)
            // commentShape.Fill.SetSolidFill(Color.LightGray);

            // Lock the shape so it cannot be edited when the sheet is protected
            commentShape.IsLocked = true;

            // Protect the worksheet to enforce the lock
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            string outputPath = "CommentShapeLocked.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}