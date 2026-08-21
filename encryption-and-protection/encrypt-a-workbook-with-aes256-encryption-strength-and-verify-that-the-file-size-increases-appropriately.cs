// Title: Encrypt an Excel workbook with AES‑256 using Aspose.Cells for .NET and check size growth
// Description: This C# example creates a workbook, saves it unprotected, records its byte size, applies a password, sets AES‑256 encryption via SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256), saves the encrypted file, and then compares the two sizes to confirm that encryption adds overhead.
// Keywords: Aspose.Cells | AES-256 encryption | C# Excel encryption | SetEncryptionOptions | StrongCryptographicProvider | Workbook password protection | XLSX file size comparison | encryption overhead | secure Excel files | Aspose.Cells .NET
// Common Searches: Aspose.Cells AES-256 encryption example | How to encrypt an Excel file with a password in C# | Verify file size increase after Excel encryption | SetEncryptionOptions StrongCryptographicProvider Aspose | C# code to compare encrypted and unencrypted XLSX sizes
// Developer Intent: Apply AES‑256 encryption to a workbook, save it, and validate that the encrypted file is larger than the original.
// Use Cases: Protect confidential spreadsheets before distribution with strong AES‑256 encryption. | Estimate storage impact of encrypted Excel reports generated programmatically. | Automate compliance‑driven logging of encryption overhead for audit trails.
// AI Prompts: Generate C# code that encrypts an Aspose.Cells workbook with AES‑256 and prints the original and encrypted file sizes. | Explain why AES‑256 encryption increases the size of an XLSX file when using Aspose.Cells. | Provide a step‑by‑step tutorial for verifying encryption overhead on Excel files with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // This C# example creates a workbook, saves it unprotected, records its byte size, applies a password, sets AES‑256 encryption via SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256), saves the encrypted file, and then compares the two sizes to confirm that encryption adds overhead.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Test");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Path for the unencrypted file
            string unencryptedPath = "Unencrypted.xlsx";
            // Save the workbook without any protection
            workbook.Save(unencryptedPath, SaveFormat.Xlsx);

            // Record the size of the unencrypted file
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // Apply password protection
            workbook.Settings.Password = "StrongPassword123";

            // Set encryption options to use AES‑256 (key length 256 bits)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Path for the encrypted file
            string encryptedPath = "Encrypted.xlsx";
            // Save the encrypted workbook
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Record the size of the encrypted file
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // Output the file sizes and verification result
            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
            Console.WriteLine($"Encryption increased size: {encryptedSize > unencryptedSize}");
        }
    }
}
