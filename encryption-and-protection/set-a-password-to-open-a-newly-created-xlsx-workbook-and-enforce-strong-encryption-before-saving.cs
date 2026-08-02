// Title: Encrypt a New XLSX Workbook with Password and AES‑128 Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with Aspose.Cells, write data, assign an opening password, apply AES‑128 strong encryption via SetEncryptionOptions, save as XLSX, and reopen it with LoadOptions to verify the protection.
// Keywords: Aspose.Cells C# | XLSX password protection | AES 128 encryption Aspose | SetEncryptionOptions | LoadOptions password | strong cryptographic provider | secure Excel file C# | encrypt workbook programmatically
// Common Searches: C# set password on Excel file Aspose.Cells | How to apply AES‑128 encryption to XLSX with Aspose | Aspose.Cells encrypt workbook example | Open password‑protected workbook using LoadOptions | Supported encryption types in Aspose.Cells
// Developer Intent: Apply an opening password and AES‑128 strong encryption to a newly created XLSX workbook before saving.
// Use Cases: Distribute confidential financial reports that require password‑based access control. | Meet regulatory compliance by encrypting generated Excel files with AES‑128. | Automate creation of per‑user encrypted workbooks in a multi‑tenant SaaS application. | Store sensitive data in Excel format on shared drives while ensuring it cannot be opened without authentication.
// AI Prompts: Generate C# code that creates an Excel workbook, protects it with a password, and uses AES‑256 encryption with Aspose.Cells. | Show how to detect whether an existing workbook is encrypted and upgrade it to a stronger encryption algorithm using Aspose.Cells. | Explain the steps to open a password‑protected XLSX file with Aspose.Cells and read a specific cell value.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook with Aspose.Cells, write data, assign an opening password, apply AES‑128 strong encryption via SetEncryptionOptions, save as XLSX, and reopen it with LoadOptions to verify the protection.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive Information");

        // Set the password required to open the workbook
        workbook.Settings.Password = "StrongPass123!";

        // Enforce strong encryption (AES 128-bit)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook as XLSX
        workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);

        // Load the workbook using the password to verify encryption
        LoadOptions loadOptions = new LoadOptions { Password = "StrongPass123!" };
        Workbook loadedWorkbook = new Workbook("EncryptedWorkbook.xlsx", loadOptions);

        // Output the loaded cell value
        Console.WriteLine("Loaded cell value: " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);
    }
}
