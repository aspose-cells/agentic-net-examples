using System;
using Aspose.Cells;

namespace AsposeCellsCommentPrivacyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIdx = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIdx];
            comment.Author = "John Doe";
            comment.Note = "Sensitive information";

            // Privacy rule: hide comments authored by users not in the allowed list
            string[] allowedAuthors = { "Admin", "ComplianceOfficer" };
            bool isAllowed = Array.Exists(allowedAuthors, a => a.Equals(comment.Author, StringComparison.OrdinalIgnoreCase));

            // Set visibility based on the privacy check
            comment.IsVisible = isAllowed; // true only for allowed authors, false otherwise

            // Optionally, hide all comments regardless of author for strict privacy
            // comment.IsVisible = false;

            // Save the workbook
            workbook.Save("CommentPrivacyDemo.xlsx");
        }
    }
}