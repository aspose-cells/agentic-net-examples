// Title: How to change an Excel comment's font color to red using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set comment text color to red | change font color of a cell comment with Aspose.Cells .NET | C# code to modify Excel comment appearance using Shape.Font.Color | apply red font to comment note in Aspose.Cells workbook | update comment shape color in generated XLSX using Aspose.Cells
// Tags: Aspose.Cells comment font color C# | Shape.Font.Color Excel comment .NET | set comment text color Aspose.Cells | modify comment appearance XLSX C# | red comment font Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentFontColor
{
    // The example creates a workbook, adds a comment to cell B2, retrieves the comment's Shape object, sets its Font.Color property to Color.Red, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "CommentFontColorDemo.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Optional sample data
                worksheet.Cells["B2"].PutValue("Sample Data");

                // Add a comment to cell B2
                int commentIndex = worksheet.Comments.Add("B2");
                Comment comment = worksheet.Comments[commentIndex];
                comment.Note = "This is a comment whose font color will be changed to red.";

                // Retrieve the shape that represents the comment
                Shape commentShape = comment.CommentShape;

                // Change the font color of the comment text to red
                commentShape.Font.Color = Color.Red;

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The comment's font color has been set to red.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
