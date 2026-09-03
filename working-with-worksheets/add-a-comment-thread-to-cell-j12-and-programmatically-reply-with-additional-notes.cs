// Title: Add a comment to cell J12 and save the workbook with Aspose.Cells for .NET (threaded comments not supported)
// AI Prompts: Using Aspose.Cells for .NET, load an existing Excel file or create a new workbook, insert a comment with the text "Initial comment thread started." into cell J12, and save the workbook as output.xlsx. | Write C# code that checks whether a workbook exists, adds a comment to cell J12, and includes a note about the current lack of threaded comment support in Aspose.Cells.
// Common Searches: Aspose.Cells C# example for adding a comment to cell J12 | How to programmatically insert a comment into a specific Excel cell using Aspose.Cells for .NET | C# load workbook, add comment, and save with Aspose.Cells | Is threaded comment functionality available in Aspose.Cells .NET? | Save Excel file after adding a comment with Aspose.Cells API
// Tags: add comment to cell Aspose.Cells C# | Aspose.Cells comment API usage | save workbook after comment Aspose.Cells | threaded comment limitation Aspose.Cells .NET | load or create workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook or creates a new one, adds a comment with initial text to cell J12, notes that threaded comments are not supported, and saves the file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet (creates one if the workbook is new).
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell J12.
            int commentIdx = sheet.Comments.Add("J12");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Initial comment thread started.";

            // NOTE: Threaded comments are not available in the current Aspose.Cells version.
            // If needed, additional replies can be added using the CommentThread API when supported.

            // Save the workbook to the specified output path.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
