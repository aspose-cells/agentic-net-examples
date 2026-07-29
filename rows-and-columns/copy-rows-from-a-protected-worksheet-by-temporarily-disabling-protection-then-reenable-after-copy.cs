// Title: Copy rows from a protected worksheet with Aspose.Cells for .NET (temporary unprotect)
// Description: C# example that loads a workbook, detects worksheet protection, temporarily unprotects the sheet, copies selected rows to another workbook, then reapplies the original protection before saving.
// Keywords: Aspose.Cells copy rows | protected worksheet C# | unprotect worksheet programmatically | Aspose.Cells temporary disable protection | CopyRows method Aspose.Cells | Excel row copy password protected | C# Aspose.Cells worksheet protection | copy rows between workbooks | Aspose.Cells protection API
// Common Searches: Aspose.Cells copy rows from protected sheet | How to unprotect a worksheet with Aspose.Cells in C# | Copy rows between Excel workbooks while keeping protection | C# copy rows from password‑protected Excel using Aspose | Temporarily disable worksheet protection Aspose.Cells
// Developer Intent: Copy specific rows from a worksheet that is locked, temporarily lift the protection, perform the copy, and restore the original protection settings.
// Use Cases: Extract header rows from a secured template and paste them into a new report workbook while keeping the template locked. | Migrate data rows from a password‑protected source sheet to a summary sheet, ensuring the source remains protected after the operation. | Create a partial copy of a protected sheet for distribution, applying the same protection type and password to the new workbook.
// AI Prompts: Generate C# code using Aspose.Cells to copy rows 5‑10 from a protected worksheet, handling an unknown password and re‑applying protection on both source and destination sheets. | Show how to detect worksheet protection, temporarily unprotect it, copy a range of rows to another workbook, and then protect both worksheets with the same ProtectionType. | Explain the steps to temporarily disable worksheet protection in Aspose.Cells, perform row copy operations, and restore the original protection without losing the password.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsProtectedRowCopyDemo
{
    // C# example that loads a workbook, detects worksheet protection, temporarily unprotects the sheet, copies selected rows to another workbook, then reapplies the original protection before saving.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceProtected.xlsx";
                const string sourceModifiedPath = "SourceProtected_Modified.xlsx";
                const string destPath = "DestinationCopiedRows.xlsx";

                // Load source workbook; create a placeholder if the file does not exist
                Workbook sourceWorkbook;
                if (File.Exists(sourcePath))
                {
                    sourceWorkbook = new Workbook(sourcePath);
                }
                else
                {
                    sourceWorkbook = new Workbook();
                    sourceWorkbook.Save(sourcePath);
                }

                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Preserve original protection settings
                bool wasProtected = sourceSheet.IsProtected;
                string password = "mypassword"; // replace with actual password if known
                ProtectionType protectionType = ProtectionType.All;

                // Attempt to unprotect; handle incorrect or missing password gracefully
                if (wasProtected)
                {
                    try
                    {
                        sourceSheet.Unprotect(password);
                    }
                    catch (CellsException)
                    {
                        try
                        {
                            sourceSheet.Unprotect();
                        }
                        catch (CellsException ex)
                        {
                            Console.WriteLine($"Unable to unprotect worksheet: {ex.Message}");
                        }
                    }
                }

                // Create destination workbook
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define rows to copy (example: rows 0‑2)
                int sourceStartRow = 0;
                int rowsToCopy = 3;
                int destStartRow = 0;

                // Copy rows from source to destination
                destSheet.Cells.CopyRows(sourceSheet.Cells, sourceStartRow, destStartRow, rowsToCopy);

                // Re‑apply protection to the source sheet if it was originally protected
                if (wasProtected)
                {
                    // oldPassword is not required when setting a new password; pass null
                    sourceSheet.Protect(protectionType, password, null);
                }

                // Optionally protect the destination sheet similarly
                destSheet.Protect(protectionType, password, null);

                // Save both workbooks
                sourceWorkbook.Save(sourceModifiedPath);
                destWorkbook.Save(destPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
