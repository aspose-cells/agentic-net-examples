using System;
using Aspose.Cells;

namespace CommentPrivacyDemo
{
    // Author: Aspose.Cells .NET example – ensures comment visibility respects privacy.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell B2
            int commentIdx = sheet.Comments.Add("B2");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Sensitive information – hide from unauthorized users";

            // Enforce privacy: make the comment invisible by default
            comment.IsVisible = false;

            // Optionally, you could toggle visibility based on a user role check
            // bool userHasPermission = CheckUserPermission();
            // comment.IsVisible = userHasPermission;

            // Save the workbook (save rule)
            workbook.Save("CommentPrivacyDemo.xlsx");
        }

        // Placeholder for a real permission check implementation
        // static bool CheckUserPermission()
        // {
        //     // Implement your privacy logic here
        //     return false;
        // }
    }
}