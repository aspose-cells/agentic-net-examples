// Title: Encrypt an Excel file with 256‑bit Microsoft Strong Cryptographic Provider using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook with Aspose.Cells, assign a password, enable 256‑bit encryption via the Microsoft Strong Cryptographic Provider (SetEncryptionOptions), and save the result as a protected XLSX file.
// Keywords: Aspose.Cells | C# | 256-bit encryption | Microsoft Strong Cryptographic Provider | SetEncryptionOptions | Workbook password | Encrypt XLSX | StrongCryptographicProvider | .NET Excel protection
// Common Searches: Aspose.Cells 256 bit encryption C# | How to use Microsoft Strong Cryptographic Provider with Aspose.Cells | Set password and encrypt Excel workbook .NET | Encrypt XLSX file using Aspose.Cells SetEncryptionOptions | Strong encryption for Excel files in C#
// Developer Intent: Apply a 256‑bit password‑protected encryption scheme to an Excel workbook before persisting it.
// Use Cases: Secure confidential financial statements before distribution. | Protect exported data from a web service by saving it as an encrypted workbook. | Achieve compliance with data‑privacy regulations that require strong file encryption.
// AI Prompts: Provide C# code that encrypts an Excel workbook with a 256‑bit password using Aspose.Cells and the Microsoft Strong Cryptographic Provider. | Show an example of setting a workbook password and enabling strong encryption before saving with Aspose.Cells for .NET. | Explain the differences between Aspose.Cells EncryptionType.StrongCryptographicProvider and other encryption options.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook with Aspose.Cells, assign a password, enable 256‑bit encryption via the Microsoft Strong Cryptographic Provider (SetEncryptionOptions), and save the result as a protected XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // (Optional) Add some data to demonstrate encryption
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("This workbook is encrypted with strong encryption.");

        // Assign a password for encryption
        workbook.Settings.Password = "MyStrong256BitPassword!";

        // Apply strong encryption (Microsoft Strong Cryptographic Provider) with a 256‑bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        workbook.Save("StrongEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
