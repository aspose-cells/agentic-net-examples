// Title: Copy rows from a password‑protected worksheet using Aspose.Cells for .NET (unprotect → copy → protect)
// Description: Demonstrates how to temporarily unprotect a worksheet, copy specific rows to another workbook with Cells.CopyRows, and then re‑apply protection. The example creates a source workbook, protects it, copies the first three rows to a destination workbook, optionally protects the destination, and saves both files.
// Keywords: Aspose.Cells copy rows protected sheet | C# unprotect worksheet Aspose.Cells | Cells.CopyRows example | protect worksheet after copy .NET | copy rows between workbooks Aspose | temporary worksheet unprotect | Aspose.Cells password protection
// Common Searches: How to copy rows from a password protected worksheet in Aspose.Cells | Aspose.Cells unprotect worksheet temporarily to copy data | Copy rows between workbooks while keeping sheet protection | Re‑apply worksheet protection after copying rows Aspose.Cells | C# Aspose.Cells copy rows from protected sheet
// Developer Intent: Copy selected rows from a protected source worksheet to another workbook by disabling protection only for the copy operation and restoring it afterward.
// Use Cases: Extract specific rows from a secured template and insert them into a report without exposing the original password. | Archive rows from a protected data sheet while keeping the source sheet locked for compliance. | Generate a shareable workbook containing only certain rows from a protected sheet, ensuring the original workbook remains protected.
// AI Prompts: Write C# code with Aspose.Cells that unprotects a worksheet, copies rows 1‑3 to a new workbook using Cells.CopyRows, then re‑protects both sheets. | Show how to copy rows 5‑10 from a password‑protected sheet to another workbook and keep the source sheet protected after the operation. | Explain how to copy rows from a protected worksheet without losing formatting or formulas using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to temporarily unprotect a worksheet, copy specific rows to another workbook with Cells.CopyRows, and then re‑apply protection. The example creates a source workbook, protects it, copies the first three rows to a destination workbook, optionally protects the destination, and saves both files.
class CopyRowsFromProtectedWorksheet
{
    static void Main()
    {
        // Create a source workbook and add some data
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Name = "Source";
        sourceSheet.Cells["A1"].PutValue("Row 1");
        sourceSheet.Cells["A2"].PutValue("Row 2");
        sourceSheet.Cells["A3"].PutValue("Row 3");

        // Protect the source worksheet with a password
        string password = "myPassword";
        sourceSheet.Protect(ProtectionType.All, password, null);

        // Create a destination workbook where rows will be copied to
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
        destinationSheet.Name = "Destination";

        // Temporarily unprotect the source worksheet to allow copying
        sourceSheet.Unprotect(password);

        // Copy the first three rows (indices 0,1,2) from source to destination starting at row 0
        destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, 3);

        // Re‑apply protection to the source worksheet
        sourceSheet.Protect(ProtectionType.All, password, null);

        // Optionally protect the destination worksheet as well
        destinationSheet.Protect(ProtectionType.All, password, null);

        // Save both workbooks
        sourceWorkbook.Save("SourceProtected.xlsx");
        destinationWorkbook.Save("DestinationWithCopiedRows.xlsx");
    }
}
