// Title: Copy Excel comment shape background and font colors from one cell to another with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells copy comment shape background color from A1 to B2 in C# | How to duplicate Excel comment font color programmatically using Aspose.Cells | C# code to transfer comment formatting between cells with Aspose.Cells | Set target comment fill color based on source comment Aspose.Cells .NET | Copy Excel cell comment style to another cell using Aspose.Cells library
// Tags: comment shape formatting copy Aspose.Cells | transfer comment fill color C# | comment font color transfer Aspose.Cells | Excel comment style replication .NET | copy comment shape properties between cells | Aspose.Cells comment formatting automation

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, ensures a source comment in A1 with specific fill and font colors, creates a target comment in B2 if missing, then copies the source comment's background (FillFormat.ForeColor) and font color to the target comment, and finally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";

                // Add a sample comment to A1 so the copy operation has a source
                int srcIdx = ws.Comments.Add("A1");
                Comment srcComment = ws.Comments[srcIdx];
                srcComment.CommentShape.Text = "Sample comment";
                srcComment.CommentShape.FillFormat.ForeColor = Color.Yellow;
                srcComment.CommentShape.Font.Color = Color.Blue;
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve source and target comments
            Comment sourceComment = sheet.Comments["A1"];
            Comment targetComment = sheet.Comments["B2"];

            // If the target comment does not exist, create it
            if (targetComment == null)
            {
                int targetIdx = sheet.Comments.Add("B2");
                targetComment = sheet.Comments[targetIdx];
            }

            // Copy formatting only when source comment exists
            if (sourceComment != null && targetComment != null)
            {
                // Copy background (fill) color
                targetComment.CommentShape.FillFormat.ForeColor = sourceComment.CommentShape.FillFormat.ForeColor;

                // Copy font color
                targetComment.CommentShape.Font.Color = sourceComment.CommentShape.Font.Color;
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
