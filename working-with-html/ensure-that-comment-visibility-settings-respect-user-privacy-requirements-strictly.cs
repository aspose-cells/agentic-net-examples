// Title: Control Excel Comment Visibility by User Authorization with Aspise.Cells for .NET
// Description: Demonstrates how to add a comment to a cell, evaluate a user‑authorization check, and set the comment's IsVisible property so that only authorized users can see the comment before saving the workbook.
// Keywords: Aspose.Cells comment visibility | Excel comment privacy .NET | IsVisible property Aspose.Cells | conditional comment display | user authorization Excel comment
// Common Searches: hide Excel comments for unauthorized users Aspose.Cells | set comment visibility based on role .NET | Aspose.Cells conditional comment display example | Excel comment privacy implementation C# | control comment visibility with user authentication
// Developer Intent: Show how to make an Excel comment visible only when a user passes an authorization test by toggling the comment's IsVisible flag.
// Use Cases: Add confidential notes to cells that only administrators can view. | Integrate custom security logic (e.g., Active Directory groups) to control comment exposure in generated reports. | Create workbooks where sensitive comments are hidden from external partners or unprivileged users.
// AI Prompts: Generate C# code using Aspose.Cells that shows a comment only for members of a specific security group. | Explain how to retrieve user roles from Azure AD and set comment.IsVisible accordingly in an Aspose.Cells workbook. | Describe how to preserve comment visibility settings when opening, editing, and re‑saving an existing Excel file.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentPrivacyDemo
{
    // Demonstrates how to add a comment to a cell, evaluate a user‑authorization check, and set the comment's IsVisible property so that only authorized users can see the comment before saving the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "Sensitive information";

            // Determine if the current user is authorized to view the comment
            bool isUserAuthorized = CheckUserAuthorization();

            // Enforce privacy: make the comment visible only for authorized users
            comment.IsVisible = isUserAuthorized;

            // Save the workbook
            workbook.Save("CommentPrivacyDemo.xlsx");
        }

        // Simulated method for user authorization logic
        private static bool CheckUserAuthorization()
        {
            // Replace with real authorization check as needed.
            // Returning false hides the comment for all users in this demo.
            return false;
        }
    }
}
