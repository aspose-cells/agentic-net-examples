// Title: C# – Save an XLSX workbook with structure protection and password encryption using Aspose.Cells
// Description: Creates a new Workbook, writes a value to cell A1, applies structure protection, sets an opening password (encryption), optionally selects the encryption algorithm and key size, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# protect workbook structure | Aspose.Cells set open password | Excel encryption Aspose.Cells | save XLSX with encryption | EncryptionType | ProtectionType.Structure | StrongCryptographicProvider | 128-bit key | .NET
// Common Searches: How to protect Excel workbook structure with Aspose.Cells C# | Set opening password for XLSX using Aspose.Cells .NET | Encrypt Excel file with 128‑bit key in C# | Save protected and encrypted workbook with Aspose.Cells | Aspose.Cells encryption options example
// Developer Intent: The developer needs C# code that generates an XLSX file which is both structure‑protected and encrypted with a password, with optional control over the encryption algorithm and key length.
// Use Cases: Distribute a template where sheet order cannot be changed and the file requires a password to open. | Create confidential financial reports that must be encrypted with a 128‑bit key before sharing. | Automate export of sensitive data, applying workbook protection and encryption in a single step.
// AI Prompts: Provide C# Aspose.Cells code to protect workbook structure, set an opening password, choose 256‑bit AES encryption, and save as XLSX. | Write a snippet that demonstrates workbook protection, custom encryption settings, and file saving with Aspose.Cells for .NET. | Explain how to apply structure protection and password encryption to an Excel file using Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Creates a new Workbook, writes a value to cell A1, applies structure protection, sets an opening password (encryption), optionally selects the encryption algorithm and key size, and saves the file as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Protected and Encrypted");

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "protectPwd");

        // Set a password required to open the workbook (encryption)
        workbook.Settings.Password = "openPwd";

        // Optional: specify encryption algorithm and key length
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the workbook in XLSX format
        workbook.Save("ProtectedEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
