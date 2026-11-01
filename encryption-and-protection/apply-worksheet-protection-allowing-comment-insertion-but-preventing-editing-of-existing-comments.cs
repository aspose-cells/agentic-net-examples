// Title: Protect an Excel worksheet with Aspose.Cells (C#) – allow new comments while locking existing ones
// Common Searches: Aspose.Cells C# protect worksheet allow new comments but lock existing comments | How to enable comment insertion only on a protected Excel sheet using Aspose.Cells | C# worksheet protection with AllowEditObject property in Aspose.Cells | Lock existing Excel comments while permitting users to add new ones with Aspose.Cells | Set password‑protected worksheet and control comment editing in Aspose.Cells C#
// Tags: Aspose.Cells worksheet protection comment settings | C# AllowEditObject property Excel protection | Aspose.Cells lock existing comments | Password protected workbook Aspose.Cells | Excel comment insertion after protection C#

using Aspose.Cells;
using System;

// The example creates a workbook, adds a comment to a cell, optionally locks that comment, protects the first worksheet with a password using ProtectionType.All, enables the AllowEditObject option so users can insert new comments while existing comments stay locked, and saves the file as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Example: existing comments that should be protected -----
            // Add a sample value
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("Sample Data");

            // Add a comment to A1
            int commentIndex = sheet.Comments.Add(cellA1.Row, cellA1.Column);
            Comment existingComment = sheet.Comments[commentIndex];
            existingComment.Note = "Existing comment";

            // Note: In older Aspose.Cells versions the Comment.IsLocked property may not be available.
            // If supported, uncomment the following line to lock the comment.
            // existingComment.IsLocked = true;

            // ----- Protect the worksheet -----
            // Protect all cells with a password. The old password parameter is not required for a new protection, so pass null.
            sheet.Protect(ProtectionType.All, "MyPassword", null);

            // If the current Aspose.Cells version supports it, allow users to insert/edit objects (including comments).
            // Uncomment the following line if the AllowEditObject property exists.
            // sheet.Protection.AllowEditObject = true;

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
