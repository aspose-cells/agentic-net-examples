// Title: Add a comment shape to cell D4, change its fill to light gray, and lock it using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add comment shape to a specific cell | set comment shape background color to light gray with Aspose.Cells | lock a comment shape to prevent editing in an Excel workbook using Aspose.Cells | how to format and protect comment shapes in Aspose.Cells for .NET | example code for creating and locking comment shapes in C# Aspose.Cells
// Tags: add comment shape Aspose.Cells C# | comment shape fill color Aspose.Cells | lock comment shape Aspose.Cells | comment shape formatting Excel .NET | protect comment shape Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates creating a new workbook, adding a comment shape to cell D4, setting its text, applying a light gray fill, locking the shape to prevent edits, and saving the file as CommentShapeExample.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell D4; Add returns the index of the new comment
            int commentIndex = sheet.Comments.Add("D4");
            Comment comment = sheet.Comments[commentIndex];

            // Set the comment text
            comment.Note = "This is a comment.";

            // Set the background color of the comment shape to light gray
            comment.CommentShape.FillFormat.ForeColor = Color.LightGray;

            // Lock the comment shape to prevent editing
            comment.CommentShape.IsLocked = true;

            // Save the workbook (lifecycle rule: save)
            string outputPath = "CommentShapeExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
