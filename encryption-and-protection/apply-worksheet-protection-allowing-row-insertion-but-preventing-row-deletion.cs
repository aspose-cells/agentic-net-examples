// Title: Aspose.Cells .NET – Protect worksheet, allow row insertion and block row deletion (password protected)
// Description: Demonstrates how to create a workbook, enable worksheet protection, set AllowInsertingRow = true, AllowDeletingRow = false, apply a password, protect all sheet features with ProtectionType.All, and save the file.
// Keywords: Aspose.Cells worksheet protection | AllowInsertingRow C# | AllowDeletingRow C# | protect sheet password Aspose.Cells | insert rows only Aspose.Cells | prevent row deletion Aspose.Cells | ProtectionType.All example | C# Aspose.Cells security
// Common Searches: Aspose.Cells protect sheet allow insert rows only | C# prevent row deletion with worksheet protection Aspose.Cells | How to set AllowInsertingRow and AllowDeletingRow in Aspose.Cells | Password‑protected worksheet that blocks row removal Aspose.Cells .NET | Enable row addition but disable row removal in Excel via Aspose.Cells
// Developer Intent: Create a protected worksheet where users can add rows but cannot delete existing rows.
// Use Cases: Data‑entry templates that let staff append records while preserving previous entries. | Financial or inventory reports that require expanding sections without risking accidental row loss. | Shared spreadsheets where contributors may insert new data rows but must not alter historic rows.
// AI Prompts: Generate C# code using Aspose.Cells to protect a worksheet, enable row insertion, disable row deletion, and set a password. | Explain the purpose of AllowInsertingRow and AllowDeletingRow flags in Aspose.Cells worksheet protection and how to apply ProtectionType.All.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, enable worksheet protection, set AllowInsertingRow = true, AllowDeletingRow = false, apply a password, protect all sheet features with ProtectionType.All, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = worksheet.Protection;

        // Allow inserting rows while the sheet is protected
        protection.AllowInsertingRow = true;

        // Disallow deleting rows while the sheet is protected
        protection.AllowDeletingRow = false;

        // Optional: set a password for the protection
        protection.Password = "password123";

        // Apply protection to the worksheet (protect all aspects)
        worksheet.Protect(ProtectionType.All);

        // Save the workbook to a file
        workbook.Save("RowProtectionDemo.xlsx");
    }
}
