// Title: C# – Apply Write‑Protection with Author and Audit Comment in Aspose.Cells
// Description: Demonstrates how to create an Aspose.Cells workbook, add a cell value, insert an audit comment with a specific author, configure write‑protection (author, password, RecommendReadOnly), save the file, and reload it to verify protection settings and comment details.
// Keywords: Aspose.Cells write protection C# | set workbook author Aspose.Cells | add comment Aspose.Cells | recommend read‑only Excel | password protect Excel file .NET | audit trail workbook protection | WriteProtection API Aspose.Cells | C# Excel security example
// Common Searches: Aspose.Cells set write protection author | C# add audit comment to Excel workbook | How to recommend read‑only mode with Aspose.Cells | Verify write protection status after saving workbook | Protect Excel file with password using Aspose.Cells .NET
// Developer Intent: Enable write‑protection on a workbook while recording the protecting author and an audit comment for traceability.
// Use Cases: Lock a financial report after an auditor adds a review comment, allowing only authorized users with a password to edit. | Distribute a compliance template that includes a reviewer’s note and is opened as read‑only until the correct password is entered. | Maintain an audit trail in a shared spreadsheet by embedding the protection author and comment, then enforcing write‑protection.
// AI Prompts: Generate code that reads the write‑protection password from appsettings.json and applies it to an Aspose.Cells workbook. | Show how to programmatically check if a loaded workbook is write‑protected before permitting edits. | Explain how to set the audit comment author dynamically based on the current Windows user.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Demonstrates how to create an Aspose.Cells workbook, add a cell value, insert an audit comment with a specific author, configure write‑protection (author, password, RecommendReadOnly), save the file, and reload it to verify protection settings and comment details.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Audit Data");

            // Add a comment for audit tracking
            int commentIndex = sheet.Comments.Add("A1");
            sheet.Comments[commentIndex].Note = "Reviewed and approved.";
            sheet.Comments[commentIndex].Author = "Audit Team";

            // ---------- Configure write‑protection ----------
            // Set the author of the protection (audit author)
            workbook.Settings.WriteProtection.Author = "Audit Team";

            // Set a password required to modify the workbook
            workbook.Settings.WriteProtection.Password = "SecurePass123";

            // Recommend opening the file as read‑only
            workbook.Settings.WriteProtection.RecommendReadOnly = true;

            // ---------- Save the protected workbook ----------
            string outputPath = "WriteProtectedAuditWorkbook.xlsx";
            workbook.Save(outputPath);                              // save

            // ---------- Load the workbook to verify protection ----------
            Workbook loadedWorkbook = new Workbook(outputPath);     // load
            WriteProtection wp = loadedWorkbook.Settings.WriteProtection;

            Console.WriteLine("Write Protection Author: " + wp.Author);
            Console.WriteLine("Is Write Protected: " + wp.IsWriteProtected);
            Console.WriteLine("Recommend Read‑Only: " + wp.RecommendReadOnly);

            // Verify the audit comment
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            if (loadedSheet.Comments.Count > 0)
            {
                var comment = loadedSheet.Comments[0];
                Console.WriteLine("Comment Author: " + comment.Author);
                Console.WriteLine("Comment Text: " + comment.Note);
            }
        }
    }
}
