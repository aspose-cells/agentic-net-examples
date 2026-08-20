// Title: C# – Save Aspose.Cells Workbook as XLSX with Structure Protection and Password Encryption
// Description: Creates a new Workbook, applies structure protection, sets an opening password, optionally configures 128‑bit strong encryption, and saves the file in XLSX format using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# protect workbook structure | Aspose.Cells set opening password | Aspose.Cells encrypt Excel file | Aspose.Cells SaveFormat.Xlsx | Aspose.Cells SetEncryptionOptions
// Common Searches: Aspose.Cells protect workbook structure and encrypt file | C# save Excel with password using Aspose.Cells | How to set opening password for XLSX in Aspose.Cells | Strong encryption options for Aspose.Cells workbook | Save protected workbook as XLSX with Aspose.Cells .NET
// Developer Intent: Generate an XLSX workbook that is both structure‑protected and encrypted with a password, using Aspose.Cells in C#.
// Use Cases: Secure financial reports so users cannot add, delete, or rename sheets without a password. | Distribute Excel templates that allow data entry but block structural changes, while keeping the file confidential. | Automate creation of confidential spreadsheets in a web service, ensuring they are saved with strong encryption before delivery.
// AI Prompts: Write C# code with Aspose.Cells to protect workbook structure (password: protectPwd), set an opening password (openPwd), enable 128‑bit strong encryption, and save as ProtectedEncryptedWorkbook.xlsx. | Show how to modify the example to use AES‑256 encryption instead of 128‑bit in Aspose.Cells. | Provide steps to programmatically verify that a saved XLSX file requires the opening password and that structure protection is active.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionExample
{
    // Creates a new Workbook, applies structure protection, sets an opening password, optionally configures 128‑bit strong encryption, and saves the file in XLSX format using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the standard creation rule)
            Workbook workbook = new Workbook();

            // Apply structure protection with a password
            workbook.Protect(ProtectionType.Structure, "protectPwd");

            // Set a password required to open the workbook (encryption)
            workbook.Settings.Password = "openPwd";

            // Optionally set stronger encryption (uses SetEncryptionOptions rule)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the protected and encrypted workbook in XLSX format (uses the standard save rule)
            workbook.Save("ProtectedEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
