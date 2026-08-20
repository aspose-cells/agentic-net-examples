// Title: Protect Excel Workbook Structure with Aspose.Cells (C#) and Block Sheet Deletion
// Description: Demonstrates how to protect a workbook's structure using Aspose.Cells for .NET, attempt to delete a worksheet, catch the resulting exception, and save the protected file.
// Keywords: Aspose.Cells protect workbook structure | C# workbook structure protection | prevent worksheet deletion Aspose.Cells | Excel file structure lock | Aspose.Cells exception on sheet removal | protect Excel template C#
// Common Searches: how to lock workbook structure using Aspose.Cells | C# code to stop sheet deletion in Excel | Aspose.Cells protect structure example | exception when removing protected worksheet Aspose.Cells | save protected Excel file with Aspose.Cells
// Developer Intent: Apply password‑protected structure protection to an Excel workbook and verify that worksheet removal is disallowed.
// Use Cases: Distribute a template where users cannot add, delete, or reorder sheets. | Create a read‑only report that preserves the original sheet layout. | Programmatically confirm that structure protection is active by handling the deletion exception.
// AI Prompts: Generate C# code with Aspose.Cells to protect only the workbook structure and show how to catch the exception when a protected sheet is deleted. | Suggest a method to check workbook.IsProtected before attempting to remove a worksheet using Aspose.Cells. | Explain how to unprotect the workbook structure, delete a sheet, and then re‑apply protection in C#.

using System;
using Aspose.Cells;

// Demonstrates how to protect a workbook's structure using Aspose.Cells for .NET, attempt to delete a worksheet, catch the resulting exception, and save the protected file.
class ProtectWorkbookStructureDemo
{
    static void Main()
    {
        // Create a new workbook and add extra worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "pwd123");
        Console.WriteLine("Workbook structure protected: " + workbook.Settings.IsProtected);

        // Attempt to delete a worksheet (should be blocked)
        try
        {
            // Try to remove the second worksheet (index 1)
            workbook.Worksheets.RemoveAt(1);
            Console.WriteLine("Worksheet removed (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Deletion blocked as expected: " + ex.Message);
        }

        // Save the workbook
        workbook.Save("ProtectedStructureWorkbook.xlsx");
    }
}
