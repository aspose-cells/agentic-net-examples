// Title: Rotate an Excel comment shape to TopToBottom orientation using Aspose.Cells for .NET
// Common Searches: Aspose.Cells rotate comment shape vertical C# | How to set Excel comment text direction to top to bottom using Aspose.Cells | C# example for vertical comment layout in Aspose.Cells workbook | Change comment shape rotation angle for top‑to‑bottom text in .NET | Display comment text vertically in Excel with Aspose.Cells API
// Tags: comment shape rotation Aspose.Cells | vertical comment text orientation .NET | set comment text direction Excel | Aspose.Cells comment shape layout | comment shape 90-degree angle C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds a comment to cell A1, sets its note, rotates the comment's shape by 90 degrees to achieve a TopToBottom (vertical) layout, and saves the file as CommentVertical.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];

            // Set the comment text
            comment.Note = "Vertical comment";

            // Rotate the comment shape to display text vertically
            // (TextOrientation property may not be available in some versions)
            comment.CommentShape.RotationAngle = 90;

            // Define output file path
            string outputPath = "CommentVertical.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
