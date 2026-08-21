// Title: C# – Protect Workbook Structure, Lock All Sheets, and Encrypt XLSX with Aspose.Cells
// Description: Creates a new workbook, applies structure protection with a password, locks every worksheet, sets a file‑open password for encryption, and saves the result as a password‑protected XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells protect workbook structure | encrypt Excel file C# | lock all worksheets Aspose.Cells | file password Aspose.Cells .NET | save encrypted XLSX Aspose.Cells | Aspose.Cells workbook protection example | C# Excel encryption Aspose
// Common Searches: how to protect workbook structure and encrypt xlsx with Aspose.Cells | Aspose.Cells lock all sheets and set file password C# | encrypt Excel workbook while protecting sheets using .NET | Aspose.Cells protect workbook and set open password example | C# code to apply structure protection and file encryption in Excel
// Developer Intent: Apply structure protection, sheet‑level locking, and file‑level encryption to an Excel workbook and save it as a password‑protected XLSX.
// Use Cases: Distribute confidential financial reports that require a password to open and cannot be edited without sheet passwords. | Generate template workbooks for partners where the layout is locked and the file is encrypted for secure transmission. | Automate creation of audit‑ready spreadsheets that enforce both workbook‑level and worksheet‑level security.
// AI Prompts: Show C# code that protects a workbook's structure, locks every worksheet, sets an open‑file password, and saves an encrypted XLSX with Aspose.Cells. | Explain how to use different passwords for workbook structure protection and file encryption in Aspose.Cells for .NET. | Provide a step‑by‑step Aspose.Cells example that creates a workbook, applies full sheet protection, encrypts the file, and saves it as XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Creates a new workbook, applies structure protection with a password, locks every worksheet, sets a file‑open password for encryption, and saves the result as a password‑protected XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Example data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "WorkbookStructurePwd");

            // Protect each worksheet (lock all cells, objects, etc.)
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.Protect(ProtectionType.All);
            }

            // Set a password to encrypt the workbook file (required to open)
            workbook.Settings.Password = "OpenFilePwd";

            // Save the encrypted and protected workbook as XLSX
            workbook.Save("ProtectedEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
