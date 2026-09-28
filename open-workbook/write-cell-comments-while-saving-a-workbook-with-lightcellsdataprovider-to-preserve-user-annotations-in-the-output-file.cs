// Title: Add a comment to cell A1 and save the workbook as XLSX while preserving annotations with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new workbook, inserts a comment into cell A1, creates the target folder if it does not exist, and saves the file as an XLSX keeping the comment intact. | Provide a concise Aspose.Cells snippet that attaches a user note to a specific cell and exports the workbook without losing the annotation.
// Common Searches: how to add a comment to a cell using Aspose.Cells C# and keep it after saving | saving an Excel file with comments using Aspose.Cells .NET | ensure output directory exists before calling Workbook.Save in Aspose.Cells | preserve cell annotations when exporting to XLSX with Aspose.Cells for .NET | Aspose.Cells add comment to A1 example C#
// Tags: Aspose.Cells add cell comment C# | Workbook.Save preserve comments XLSX | ensure output directory Aspose.Cells | cell annotation handling Aspose.Cells | export workbook with comments .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, adds a comment "User annotation for cell A1" to cell A1, ensures the destination directory exists, and saves the workbook as output.xlsx in XLSX format, with the comment retained automatically.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            // The Comments collection is zero‑based; Add returns the index of the new comment
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "User annotation for cell A1";

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (comments are preserved by default)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
