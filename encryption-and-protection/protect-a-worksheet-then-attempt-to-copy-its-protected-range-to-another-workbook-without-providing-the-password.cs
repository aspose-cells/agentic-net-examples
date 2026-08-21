// Title: Attempt to copy rows from a password‑protected worksheet range using Aspose.Cells for .NET
// Description: The example creates a workbook, fills cells A1:B2, defines a protected range with a password, protects the sheet, and then tries to copy those rows to another workbook via Cells.CopyRows. Because the source range is locked, the copy throws an exception, which is caught and displayed.
// Keywords: Aspose.Cells copy protected range | C# copy rows password protected worksheet | Aspose.Cells Cells.CopyRows exception | protected range copy without password | Aspose.Cells worksheet protection
// Common Searches: Aspose.Cells copy rows from a locked range | How to copy data from a password‑protected worksheet in .NET | Exception when copying protected range with Aspose.Cells | Copy rows from protected sheet without password Aspose.Cells
// Developer Intent: Show that copying rows from a worksheet containing a password‑protected range fails when the password is not supplied, and demonstrate proper exception handling.
// Use Cases: Verify that protected ranges prevent unauthorized data copying. | Capture and log the exception raised by Cells.CopyRows on a locked range. | Inform users that the correct range password is required for copy operations.
// AI Prompts: Generate C# code that copies rows from a protected worksheet after providing the range password using Aspose.Cells. | Explain how to retrieve and remove a protected range password programmatically before copying data. | Show how to catch the specific exception thrown when copying a password‑protected range without a password.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsProtectedRangeCopyDemo
{
    // The example creates a workbook, fills cells A1:B2, defines a protected range with a password, protects the sheet, and then tries to copy those rows to another workbook via Cells.CopyRows. Because the source range is locked, the copy throws an exception, which is caught and displayed.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook ----------
                Workbook srcWb = new Workbook();
                Worksheet srcSheet = srcWb.Worksheets[0];

                // Fill some data in the range A1:B2
                srcSheet.Cells["A1"].PutValue("Item");
                srcSheet.Cells["B1"].PutValue("Quantity");
                srcSheet.Cells["A2"].PutValue("Apple");
                srcSheet.Cells["B2"].PutValue(10);

                // Add a protected range covering A1:B2 and set a password
                int rangeIndex = srcSheet.AllowEditRanges.Add("MyProtectedRange", 0, 0, 1, 1);
                ProtectedRange protectedRange = srcSheet.AllowEditRanges[rangeIndex];
                protectedRange.Password = "rangePwd";

                // Protect the entire worksheet (with a password for completeness)
                srcSheet.Protect(ProtectionType.All, "sheetPwd", null);

                // Save the source workbook (optional, just to visualize)
                string srcPath = "SourceProtected.xlsx";
                try
                {
                    srcWb.Save(srcPath);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save source workbook: {saveEx.Message}");
                }

                // ---------- Attempt to copy the protected range ----------
                Workbook destWb = new Workbook();
                Worksheet destSheet = destWb.Worksheets[0];

                try
                {
                    // Copy the first two rows (A1:B2) from source to destination.
                    // Correct overload: CopyRows(destCells, startRow, totalRows, destStartRow)
                    srcSheet.Cells.CopyRows(destSheet.Cells, 0, 2, 0);
                    Console.WriteLine("Copy operation succeeded (unexpected).");
                }
                catch (Exception copyEx)
                {
                    // Expected path: operation fails because the source range is password‑protected
                    Console.WriteLine("Copy operation failed as expected: " + copyEx.Message);
                }

                // Save the destination workbook to see the result
                string destPath = "DestinationAfterCopyAttempt.xlsx";
                try
                {
                    destWb.Save(destPath);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save destination workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
