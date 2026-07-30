// Title: C# – Encrypt an Excel workbook with 256‑bit Microsoft Strong Cryptographic Provider using Aspose.Cells
// Description: Creates a new Workbook, writes sample data, assigns a password, applies 256‑bit encryption via Aspose.Cells' StrongCryptographicProvider option, and saves the file as a protected XLSX.
// Keywords: Aspose.Cells | C# | StrongCryptographicProvider | 256-bit encryption | Excel workbook password | SetEncryptionOptions | Encrypt XLSX | Microsoft CryptoAPI | Data protection | Secure Excel files
// Common Searches: Aspose.Cells 256 bit encryption C# | How to use Microsoft Strong Cryptographic Provider with Aspose.Cells | Set password and encrypt Excel file Aspose.Cells | Encrypt XLSX using SetEncryptionOptions | Strong encryption for Excel workbook .NET
// Developer Intent: Apply 256‑bit Microsoft Strong Cryptographic Provider encryption with a password to an Excel workbook via Aspose.Cells.
// Use Cases: Safeguard confidential financial reports before distribution. | Comply with GDPR or HIPAA by storing sensitive spreadsheets in encrypted form. | Automate secure generation of payroll or tax worksheets in enterprise pipelines.
// AI Prompts: Write C# code that opens an existing workbook, sets a password, applies Microsoft Strong Cryptographic Provider encryption with a 256‑bit key, and saves it. | Explain the parameters of Aspose.Cells SetEncryptionOptions for strong encryption of XLSX files.

using System;
using Aspose.Cells;

// Creates a new Workbook, writes sample data, assigns a password, applies 256‑bit encryption via Aspose.Cells' StrongCryptographicProvider option, and saves the file as a protected XLSX.
class StrongEncryptionExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Information");

        // Set the password for encryption
        workbook.Settings.Password = "MyStrongPassword123!";

        // Apply strong encryption (Microsoft Strong Cryptographic Provider) with a 256‑bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        workbook.Save("StrongEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
