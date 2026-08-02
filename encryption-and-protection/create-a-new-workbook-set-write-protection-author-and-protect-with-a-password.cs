// Title: Aspose.Cells for .NET – Create a Workbook, Set Write‑Protection Author & Password, and Save
// Description: Creates a new Workbook with Aspose.Cells, assigns a write‑protection author, sets a password, optionally enables RecommendReadOnly, and saves the file as WriteProtectedWorkbook.xlsx.
// Keywords: Aspose.Cells write protection | C# set workbook password | Excel write protection author | RecommendReadOnly Aspose.Cells | protect Excel file .NET | Aspose.Cells Workbook Settings | write protection API
// Common Searches: How to set write protection author in Aspose.Cells C# | Aspose.Cells protect workbook with password .NET | Enable RecommendReadOnly flag using Aspose.Cells | Create password‑protected Excel file with Aspose.Cells | Set workbook write protection programmatically in C#
// Developer Intent: Programmatically protect a newly created Excel workbook with an author name and password, optionally recommending read‑only mode, using Aspose.Cells for .NET.
// Use Cases: Distribute read‑only templates that require a password for edits | Secure financial or audit spreadsheets while identifying the protection author | Automate generation of confidential reports that must be opened in read‑only mode unless authorized | Enforce write‑access control in multi‑user Excel workflows
// AI Prompts: Generate C# code that adds write protection with a specific author and password to an existing Aspose.Cells workbook. | Show how to modify or remove the write‑protection password of a saved Excel file using Aspose.Cells for .NET. | Explain how the RecommendReadOnly property influences Excel's opening behavior for a password‑protected workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Creates a new Workbook with Aspose.Cells, assigns a write‑protection author, sets a password, optionally enables RecommendReadOnly, and saves the file as WriteProtectedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the author for write protection
            workbook.Settings.WriteProtection.Author = "John Doe";

            // Set the password that protects the workbook from modification
            workbook.Settings.WriteProtection.Password = "password123";

            // (Optional) Recommend opening the file as read‑only
            workbook.Settings.WriteProtection.RecommendReadOnly = true;

            // Save the workbook to a file
            workbook.Save("WriteProtectedWorkbook.xlsx");
        }
    }
}
