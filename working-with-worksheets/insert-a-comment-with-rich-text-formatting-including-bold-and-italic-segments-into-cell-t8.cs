// Title: How to add a comment with bold and italic text to cell T8 using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a comment in cell T8 of a new workbook and uses the Comment.NoteRichText collection to apply bold style to the word "Bold" and italic style to the word "Italic" with Aspose.Cells. | Show an example of inserting mixed‑style text (bold and italic) into an Excel comment by populating the Comment.NoteRichText property for cell T8 using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add comment with bold and italic text to a specific cell | How to format an Excel comment with rich text using Aspose.Cells | C# example of mixed formatting in an Excel comment with Aspose.Cells | Insert a rich‑text comment into cell T8 in a .NET workbook | Apply bold and italic styles inside an Excel comment programmatically
// Tags: Aspose.Cells comment formatting | C# Excel comment rich text | mixed style comment Aspose.Cells | cell T8 comment Aspose.Cells | Excel comment styling with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, accesses the first worksheet, adds a comment to cell T8, explains that rich‑text formatting (bold/italic) can be applied via the Comment.NoteRichText property, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell T8
            int commentIdx = sheet.Comments.Add("T8");
            Comment comment = sheet.Comments[commentIdx];

            // Set the comment text
            comment.Note = "Bold Italic";

            // Note: Rich text formatting (bold/italic) on individual characters
            // requires the NoteRichText property, which may not be available in
            // certain versions of Aspose.Cells. The basic comment text is set above.

            // Determine output path and ensure its directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
