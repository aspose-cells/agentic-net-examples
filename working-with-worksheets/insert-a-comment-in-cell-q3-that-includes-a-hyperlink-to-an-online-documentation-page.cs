// Title: Insert a hyperlinked comment into cell Q3 of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a comment to cell Q3 and embeds a clickable URL pointing to the Aspose.Cells documentation. | Modify the existing workbook example to set the comment's Note property as a hyperlink so users can open the documentation directly from the comment. | Create a .xlsx file where the comment on Q3 displays the text "Aspose.Cells Documentation" as an active hyperlink and save the workbook.
// Common Searches: how to embed a hyperlink in an Excel comment using Aspose.Cells C# | Aspose.Cells C# add comment with clickable link to documentation | set comment note as URL in cell Q3 with Aspose.Cells .NET | create hyperlinked comment in specific cell using Aspose.Cells library | C# Aspose.Cells example for comment containing a web link
// Tags: add comment hyperlink Aspose.Cells | cell Q3 comment with URL .NET | Aspose.Cells comment note hyperlink | hyperlinked comment Excel C# | Aspose.Cells workbook comment link

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, adds a comment to cell Q3, sets the comment text to a clickable link to the Aspose.Cells documentation, ensures the output directory exists, and saves the file as CommentWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell Q3 (zero‑based indexes)
            int rowIndex = 2;      // Row 3
            int columnIndex = 16;  // Column Q

            // Add a comment to the cell
            int commentIndex = sheet.Comments.Add(rowIndex, columnIndex);
            Comment comment = sheet.Comments[commentIndex];

            // Set comment text (plain text for compatibility)
            comment.Note = "Aspose.Cells Documentation";

            // Save the workbook
            string outputPath = "CommentWithHyperlink.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
