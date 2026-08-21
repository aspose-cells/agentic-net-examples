// Title: Encrypt a Merged Workbook with a Password Using Aspose.Cells for .NET (C#)
// Description: Shows how to assign a password, enable AES‑128 encryption, save, and reload a merged Excel workbook with Aspose.Cells for .NET to protect confidential data.
// Keywords: Aspose.Cells | C# encrypt Excel workbook | password protect workbook | AES 128 encryption Aspose | merged workbook security | Workbook.Settings.Password | SetEncryptionOptions | LoadOptions.Password | Excel file protection .NET | secure Excel output
// Common Searches: Aspose.Cells set password for Excel file | Encrypt merged workbook C# Aspose | AES encryption Aspose.Cells example | How to protect Excel workbook with password in .NET | Load encrypted Excel with Aspose.Cells
// Developer Intent: Apply password protection and AES encryption to a merged workbook before saving to safeguard confidential information.
// Use Cases: Securely distribute merged reports to external clients or partners. | Fulfill GDPR, HIPAA, or PCI compliance by encrypting generated Excel files. | Automate a verification step that reopens the saved file with the password to confirm encryption.
// AI Prompts: Generate C# code that merges several Excel files using Aspose.Cells and then encrypts the resulting workbook with a user‑defined password and AES‑256 encryption. | Provide an example of opening an encrypted Excel file with Aspose.Cells, handling incorrect password exceptions, and extracting specific cell values. | Create a C# snippet that encrypts a workbook, writes it to a memory stream, and attaches the encrypted file to an email message.

using System;
using Aspose.Cells;

// Shows how to assign a password, enable AES‑128 encryption, save, and reload a merged Excel workbook with Aspose.Cells for .NET to protect confidential data.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook (replace with your merged workbook if already created)
        Workbook wb = new Workbook();

        // Example: add some confidential data
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Confidential data");

        // Set a password to encrypt the workbook
        wb.Settings.Password = "StrongPassword123";

        // Optional: specify stronger encryption (AES 128-bit)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        wb.Save("MergedWorkbook_Encrypted.xlsx");

        // Verify encryption by loading the workbook with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "StrongPassword123";
        Workbook loadedWb = new Workbook("MergedWorkbook_Encrypted.xlsx", loadOptions);
        Console.WriteLine("Loaded cell value: " + loadedWb.Worksheets[0].Cells["A1"].Value);
    }
}
