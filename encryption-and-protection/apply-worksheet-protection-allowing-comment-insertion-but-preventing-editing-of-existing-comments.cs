// Title: Aspose.Cells C# – Protect Worksheet, Allow New Comments, Lock Existing Comments
// Description: Creates a workbook, adds a comment to A1, locks the comment's shape, configures worksheet protection to permit inserting new comments while blocking edits to locked comments, applies password protection, and saves the file as CommentProtection.xlsx.
// Keywords: Aspose.Cells | C# worksheet protection | Excel comment lock | Allow insert comments | Password protect sheet | CommentShape IsLocked | Secure Excel template | Object editing restriction | Aspose.Cells Protect | Excel security C#
// Common Searches: Aspose.Cells protect worksheet allow comments | Lock existing comment Aspose.Cells C# | Enable comment insertion but prevent editing existing comments Excel | C# Aspose.Cells worksheet protection settings | How to lock comment shape in Aspose.Cells
// Developer Intent: Set worksheet protection so users can add new comments while existing comments remain read‑only.
// Use Cases: Distribute a financial model where pre‑filled audit notes are immutable but reviewers can add their own remarks. | Create a shared project tracker that lets team members insert feedback without altering the original guidance. | Publish a template with locked instructions while permitting collaborators to append additional comments.
// AI Prompts: Show C# code using Aspose.Cells to lock a comment shape and protect a worksheet so only new comments can be added. | Explain how Protection.AllowEditingObject and CommentShape.IsLocked work together to restrict comment editing in Excel. | Generate an Aspose.Cells example that password‑protects a sheet while allowing comment insertion.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for Shape-related properties

// Creates a workbook, adds a comment to A1, locks the comment's shape, configures worksheet protection to permit inserting new comments while blocking edits to locked comments, applies password protection, and saves the file as CommentProtection.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Add an existing comment that we want to protect
            // -------------------------------------------------
            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment existingComment = sheet.Comments[commentIndex];
            existingComment.Note = "This comment is locked and cannot be edited.";

            // Lock the comment so it cannot be edited when the sheet is protected
            // The comment is represented as a shape; lock the shape.
            if (existingComment.CommentShape != null)
            {
                existingComment.CommentShape.IsLocked = true;
            }

            // -------------------------------------------------
            // Configure worksheet protection
            // -------------------------------------------------
            // Allow users to insert new objects (comments are objects)
            // but locked objects (the existing comment) will remain uneditable
            Protection protection = sheet.Protection;
            protection.AllowEditingObject = true;          // permit insertion of new comments
            protection.AllowEditingContent = false;        // optional: prevent editing cell contents
            protection.AllowInsertingRow = true;           // optional: allow other insertions if needed

            // Protect the worksheet with a password
            sheet.Protect(ProtectionType.All, "MySecretPwd", null);

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "CommentProtection.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
