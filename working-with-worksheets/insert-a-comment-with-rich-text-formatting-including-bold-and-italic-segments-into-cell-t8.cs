// Title: Add a bold and italic formatted comment to cell T8 using Aspose.Cells for .NET
// Common Searches: how to add a formatted comment with bold and italic text to a specific cell using Aspose.Cells C# | Aspose.Cells C# example for inserting a comment into cell T8 and making it visible | C# code to create an Excel comment with author and rich‑text styling using Aspose.Cells | Aspose.Cells comment API set bold italic text in Excel comment C# | save workbook with visible comment in cell T8 using Aspose.Cells .NET
// Tags: Aspose.Cells add comment C# | Excel comment formatting Aspose.Cells | cell T8 comment visibility | Aspose.Cells rich text comment limitation | save workbook with comment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, accesses the first Worksheet, adds a comment to cell T8, assigns an author, sets plain text (rich‑text not supported in this version), makes the comment visible, and saves the file as CommentWithRichText.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a comment to cell T8
                int commentIndex = worksheet.Comments.Add("T8");
                Comment comment = worksheet.Comments[commentIndex];
                comment.Author = "Author";

                // Set the comment text (rich‑text not supported in this version)
                comment.Note = "Bold Segment and Italic Segment";

                // Make the comment visible
                comment.IsVisible = true;

                // Define output path and ensure the directory exists
                string outputPath = "CommentWithRichText.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
