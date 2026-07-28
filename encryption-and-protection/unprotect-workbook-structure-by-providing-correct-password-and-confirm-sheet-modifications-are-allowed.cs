// Title: C# – Unprotect Aspose.Cells workbook structure with password and allow sheet editing
// Description: Demonstrates how to protect a workbook's structure with a password, verify protection via IsWorkbookProtectedWithPassword, remove the protection using Workbook.Unprotect, confirm the workbook is unprotected, add a new worksheet, and save the resulting file.
// Keywords: Aspose.Cells unprotect workbook C# | remove workbook structure protection | Workbook.Unprotect password | IsWorkbookProtectedWithPassword check | add worksheet after unprotect | protect workbook structure Aspose.Cells | Excel file password protection .NET | C# Aspose.Cells protect and unprotect
// Common Searches: how to unprotect workbook structure using Aspose.Cells for .NET | remove password protection from an Aspose.Cells workbook | verify workbook protection status after Unprotect in C# | add new sheet after unprotecting an Aspose.Cells workbook | Aspose.Cells C# example for protecting and unprotecting workbook structure
// Developer Intent: Remove the workbook's structure protection with the correct password so that sheets can be edited or added.
// Use Cases: Programmatically lift structure protection before inserting, renaming, or deleting worksheets. | Validate that a workbook is no longer protected after calling Unprotect to ensure further modifications are allowed. | Create an unprotected copy of a password‑protected workbook for downstream processing or distribution.
// AI Prompts: Generate C# code that uses Aspose.Cells to unprotect a workbook structure with a known password, then adds a new worksheet and saves the file. | Write a method that logs the value of IsWorkbookProtectedWithPassword before and after calling Workbook.Unprotect, handling any exceptions. | Provide an example that catches an incorrect‑password exception when calling Workbook.Unprotect in Aspose.Cells.

using Aspose.Cells;
using System;

// Demonstrates how to protect a workbook's structure with a password, verify protection via IsWorkbookProtectedWithPassword, remove the protection using Workbook.Unprotect, confirm the workbook is unprotected, add a new worksheet, and save the resulting file.
class UnprotectWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword");

        // Confirm that the workbook is protected with a password
        Console.WriteLine("Workbook protected with password: " + workbook.IsWorkbookProtectedWithPassword);

        // Unprotect the workbook using the correct password
        workbook.Unprotect("myPassword");

        // Confirm that the workbook is no longer protected
        Console.WriteLine("Workbook protected after unprotect: " + workbook.IsWorkbookProtectedWithPassword);

        // Verify that sheet modifications are now allowed by adding a new worksheet
        int newSheetIndex = workbook.Worksheets.Add();
        workbook.Worksheets[newSheetIndex].Name = "NewSheet";

        // Save the unprotected workbook
        workbook.Save("UnprotectedWorkbook.xlsx");
    }
}
