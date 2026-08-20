// Title: Programmatically Hide Excel Cell Comments with Aspose.Cells for .NET – Enforce Strict Privacy
// Description: This example creates a workbook, adds a confidential comment to cell A1, sets the comment's IsVisible property to false, iterates through all worksheet comments to ensure they are hidden, and saves the file as CommentPrivacyDemo.xlsx. The approach guarantees that sensitive comment data remains invisible, supporting GDPR‑style privacy requirements.
// Keywords: Aspose.Cells hide comment | Excel comment visibility .NET | C# hide Excel comments | comment privacy Aspose.Cells | set IsVisible false | GDPR Excel comment protection | confidential Excel comment | suppress Excel comments programmatically
// Common Searches: how to hide Excel comments using Aspose.Cells C# | set comment.IsVisible false Aspose.Cells | make Excel cell comments invisible for privacy | Aspose.Cells hide all comments in workbook | protect sensitive comment data in Excel .NET
// Developer Intent: Hide Excel cell comments so that sensitive information is not displayed, while keeping the comment data intact in the workbook.
// Use Cases: Add a confidential note to a cell and hide it before distributing the workbook. | Batch‑process a worksheet to suppress all existing comments for a public release. | Generate reports that retain comment metadata for internal audit but prevent end‑user visibility.
// AI Prompts: Generate C# code with Aspose.Cells that hides every comment in an existing workbook while preserving the comment text. | Show how to set comment.IsVisible = false for a specific cell and verify the setting after saving the file. | Explain a strategy to enforce comment privacy across multiple worksheets using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a confidential comment to cell A1, sets the comment's IsVisible property to false, iterates through all worksheet comments to ensure they are hidden, and saves the file as CommentPrivacyDemo.xlsx. The approach guarantees that sensitive comment data remains invisible, supporting GDPR‑style privacy requirements.
    public class CommentPrivacyDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully as CommentPrivacyDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Author = "SensitiveUser";
            comment.Note = "Confidential information";

            // Enforce strict privacy by making the comment invisible
            comment.IsVisible = false;

            // Ensure all existing comments in the worksheet are also invisible
            foreach (Comment c in worksheet.Comments)
            {
                c.IsVisible = false;
            }

            // Save the workbook with the privacy settings applied
            workbook.Save("CommentPrivacyDemo.xlsx");
        }
    }
}
