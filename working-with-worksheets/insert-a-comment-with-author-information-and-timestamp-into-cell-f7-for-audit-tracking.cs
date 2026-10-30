// Title: Add an author and timestamp audit comment to cell F7 in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to add a comment with author and current date to a specific cell using Aspose.Cells C# | Aspose.Cells C# insert multi‑line audit comment into Excel cell F7 | C# code to create Excel comment for audit tracking with timestamp using Aspose.Cells | add author information to Excel comment programmatically with Aspose.Cells .NET | save workbook after adding comment to cell F7 Aspose.Cells example
// Tags: add comment to Excel cell Aspose.Cells | author timestamp audit comment Aspose.Cells | insert multi‑line comment worksheet Aspose.Cells | save workbook after adding comment Aspose.Cells | cell F7 comment insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, builds a multi‑line comment containing an author name and the current timestamp, inserts the comment into cell F7 (row 6, column 5) of the first worksheet, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Author and timestamp information
            string author = "John Doe";
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Compose the comment text
            string commentText = $"Author: {author}\nTimestamp: {timestamp}\nAudit: Record updated.";

            // Insert the comment into cell F7 (zero‑based indices: row 6, column 5)
            int commentIndex = sheet.Comments.Add(6, 5); // Add empty comment at the cell
            Comment comment = sheet.Comments[commentIndex];
            comment.Author = author;
            comment.Note = commentText; // Store the full text

            // Ensure the output directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
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
