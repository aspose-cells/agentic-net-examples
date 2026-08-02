// Title: Encrypt an Excel workbook with AES‑256 using Aspose.Cells for .NET and confirm size growth
// Description: This C# example shows how to create a workbook, save it unencrypted, apply a strong password with AES‑256 encryption via Aspose.Cells Settings and SetEncryptionOptions, save the encrypted file, compare the byte sizes to demonstrate the expected increase, and finally load the protected workbook using the password to verify successful decryption.
// Keywords: Aspose.Cells | AES-256 | Excel encryption .NET | password‑protected workbook | file size comparison | SetEncryptionOptions | StrongCryptographicProvider | C# Excel security | LoadOptions password | cross‑platform .NET
// Common Searches: Aspose.Cells AES-256 encrypt Excel file | C# encrypt workbook and check file size | How to set password protection with Aspose.Cells | Verify encrypted Excel size increase | Load password protected Excel with Aspose.Cells .NET
// Developer Intent: Encrypt a workbook with AES‑256, save it, verify the encrypted file is larger, and ensure it can be opened with the password.
// Use Cases: Protect confidential spreadsheets before distribution | Measure encryption overhead for storage planning | Automate validation of password‑protected workbooks in CI pipelines | Implement compliance‑driven data protection in .NET applications
// AI Prompts: Write C# code that uses Aspose.Cells to apply AES‑256 encryption to an existing workbook, save both unencrypted and encrypted versions, output their sizes, and then open the encrypted file with the password. | Explain how Aspose.Cells SetEncryptionOptions works, which encryption strengths are supported, and how to handle decryption errors in C#.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // This C# example shows how to create a workbook, save it unencrypted, apply a strong password with AES‑256 encryption via Aspose.Cells Settings and SetEncryptionOptions, save the encrypted file, compare the byte sizes to demonstrate the expected increase, and finally load the protected workbook using the password to verify successful decryption.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Test");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Save the workbook without encryption
            string unencryptedPath = "Unencrypted.xlsx";
            workbook.Save(unencryptedPath, SaveFormat.Xlsx);
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // Set password and AES‑256 encryption options
            workbook.Settings.Password = "StrongPassword123";
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook
            string encryptedPath = "Encrypted.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // Output file sizes to verify increase
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
            Console.WriteLine($"Size increase:        {encryptedSize - unencryptedSize} bytes");

            // Verify that the encrypted workbook can be loaded with the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "StrongPassword123";
            Workbook loadedEncrypted = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine("Encrypted workbook loaded successfully. Cell A1 value: " +
                              loadedEncrypted.Worksheets[0].Cells["A1"].StringValue);
        }
    }
}
