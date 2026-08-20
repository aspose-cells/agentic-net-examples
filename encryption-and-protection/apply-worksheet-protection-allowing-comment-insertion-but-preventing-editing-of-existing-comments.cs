// Title: Aspose.Cells C# – Protect Worksheet to Allow Adding New Comments While Blocking Existing Comment Edits
// Description: Shows how to create a workbook, insert an initial comment, and protect the worksheet with Aspose.Cells so users can add new comments but cannot edit cell values or existing comments. The example sets AllowEditingObject = true, AllowEditingContent = false, applies a password, and saves the file.
// Keywords: Aspose.Cells | worksheet protection | comment insertion | prevent comment editing | C# | AllowEditingObject | Protect method | password protection | Excel template comments | read‑only comments
// Common Searches: Aspose.Cells protect sheet allow comments | C# Aspose.Cells enable comment addition only | block editing of existing comments Aspose.Cells | allow users to add comments to protected worksheet Aspose.Cells | worksheet protection settings for comments C#
// Developer Intent: Implement worksheet protection that permits new comment creation while preventing modifications to existing comments and cell content.
// Use Cases: Distribute a financial model where analyst notes stay locked but reviewers can add their own remarks. | Create a QA checklist template with immutable issue comments, allowing testers to append additional observations. | Provide an Excel template for collaborative budgeting where the original budget assumptions are protected, yet participants can insert supplemental comments.
// AI Prompts: Generate C# code using Aspose.Cells to protect a worksheet, allowing only new comment insertion and blocking edits to existing comments. | Explain the effect of the AllowEditingObject property in Aspose.Cells protection and how it interacts with other protection flags for comment‑only editing. | Show an example of protecting an Aspose.Cells worksheet with a password, disabling cell content changes, and enabling users to add comments.

using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtection
{
    // Shows how to create a workbook, insert an initial comment, and protect the worksheet with Aspose.Cells so users can add new comments but cannot edit cell values or existing comments. The example sets AllowEditingObject = true, AllowEditingContent = false, applies a password, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a sample comment to a cell (this represents an existing comment)
            Comment existingComment = sheet.Comments[sheet.Comments.Add("A1")];
            existingComment.Note = "Existing comment";

            // Configure protection settings:
            // - Allow editing of objects (comments are objects) so users can insert new comments.
            // - Disallow editing of objects that already exist by not granting edit rights after protection.
            //   In Aspose.Cells, this is controlled by the AllowEditingObject flag.
            //   Setting it to true permits object manipulation (including comment insertion).
            //   Existing comments will remain as they are unless the user explicitly edits them.
            Protection protection = sheet.Protection;
            protection.AllowEditingObject = true;          // Allow comment insertion/editing
            protection.AllowEditingContent = false;        // Prevent editing of cell contents
            protection.AllowDeletingRow = false;           // Example of other restrictions
            protection.Password = "securePwd123";

            // Protect the worksheet with all protection types
            sheet.Protect(ProtectionType.All, protection.Password, null);

            // Save the workbook
            workbook.Save("WorksheetProtectedWithComments.xlsx");
        }
    }
}
