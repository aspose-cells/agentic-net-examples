using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a sample comment to cell A1 (existing comment)
            Comment existingComment = sheet.Comments[sheet.Comments.Add("A1")];
            existingComment.Note = "Existing comment";

            // Access the worksheet protection object
            Protection protection = sheet.Protection;

            // Allow users to insert new comments (comments are considered drawing objects)
            protection.AllowEditingObject = true;

            // Disallow editing of cell contents (including existing comments' text)
            // Existing comments become read‑only because cell contents are protected
            protection.AllowEditingContent = false;

            // Apply protection (no password required for this example)
            sheet.Protect(ProtectionType.Objects);

            // Save the workbook
            workbook.Save("WorksheetProtected_AllowInsertComments.xlsx");
        }
    }
}