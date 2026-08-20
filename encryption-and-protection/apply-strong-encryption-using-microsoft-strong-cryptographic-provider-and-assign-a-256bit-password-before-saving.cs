// Title: Encrypt Excel Workbook with 256‑bit Microsoft Strong Cryptographic Provider using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add sample data, set a password, apply 256‑bit Microsoft Strong Cryptographic Provider encryption, and save the file as .xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | strong encryption | Microsoft Strong Cryptographic Provider | 256-bit | C# workbook password | Excel encryption .NET | SetEncryptionOptions | EncryptionType.StrongCryptographicProvider
// Common Searches: Aspose.Cells 256 bit encryption C# | How to use Microsoft Strong Cryptographic Provider with Aspose.Cells | Encrypt Excel file with password in .NET | Set workbook encryption type Aspose.Cells | Save encrypted .xlsx using Aspose.Cells
// Developer Intent: Apply 256‑bit Microsoft Strong Cryptographic Provider encryption and a password to an Excel workbook before saving.
// Use Cases: Secure financial statements before distribution | Protect exported data from web services | Meet regulatory requirements for encrypted Excel files | Store confidential customer information in encrypted workbooks
// AI Prompts: Provide C# code to encrypt an Aspose.Cells workbook with a custom password using EncryptionType.StrongCryptographicProvider and a 256‑bit key. | Show how to verify that a password protects an Excel file saved with Aspose.Cells strong encryption. | Compare Aspose.Cells EncryptionType options and explain when to choose StrongCryptographicProvider. | Generate a PowerShell script that calls a .NET assembly to encrypt an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add sample data, set a password, apply 256‑bit Microsoft Strong Cryptographic Provider encryption, and save the file as .xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Optional: add some data to demonstrate the workbook content
        workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted with Strong Cryptographic Provider");

        // Assign a password (the password itself is a string; the encryption will use a 256‑bit key)
        workbook.Settings.Password = "MyStrongPassword123!";

        // Apply strong encryption (Microsoft Strong Cryptographic Provider) with a 256‑bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the workbook using the standard save method
        workbook.Save("StrongEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
