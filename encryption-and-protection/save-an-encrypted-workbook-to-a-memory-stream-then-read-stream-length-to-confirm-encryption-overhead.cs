// Title: Encrypt an Aspose.Cells Workbook to a MemoryStream and Measure Size Overhead (C#)
// Description: C# example that creates an Aspose.Cells Workbook, writes sample data, saves it to a MemoryStream, records the unencrypted length, applies a password via Workbook.Settings.Password, optionally sets strong encryption options, saves the encrypted version to another MemoryStream, and calculates the byte difference to show the encryption overhead.
// Keywords: Aspose.Cells | C# encryption | Workbook.SaveToStream | password protected workbook | encryption overhead | memory stream size | Excel encryption .NET | measure encrypted file size | Aspose.Cells encryption options
// Common Searches: C# save Aspose.Cells workbook to MemoryStream with password | How to get encrypted workbook size using Aspose.Cells | Aspose.Cells encryption overhead measurement | Set password for Excel file in memory using Aspose.Cells | Compare encrypted vs unencrypted workbook stream length
// Developer Intent: Save a workbook to a MemoryStream with password protection and compare its byte size to the unencrypted version.
// Use Cases: Compliance testing to verify expected byte increase after encryption | Generate password‑protected Excel files on the fly for web APIs | Benchmark Aspose.Cells encryption performance by measuring stream length differences | Validate file size before transmitting encrypted workbooks over a network | Implement in‑memory Excel handling without writing temporary files to disk
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, save it to a MemoryStream, apply a password, and output both unencrypted and encrypted stream lengths. | Explain how to configure strong encryption (e.g., AES‑128) for an Aspose.Cells workbook when saving to a stream. | Provide a reusable method that returns the encryption overhead in bytes for any given workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // C# example that creates an Aspose.Cells Workbook, writes sample data, saves it to a MemoryStream, records the unencrypted length, applies a password via Workbook.Settings.Password, optionally sets strong encryption options, saves the encrypted version to another MemoryStream, and calculates the byte difference to show the encryption overhead.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue(123);

            // Save the workbook without encryption to a memory stream
            MemoryStream unencryptedStream = workbook.SaveToStream();
            long unencryptedLength = unencryptedStream.Length;
            Console.WriteLine($"Unencrypted stream length: {unencryptedLength} bytes");

            // Apply password protection (encryption)
            workbook.Settings.Password = "myPassword";
            // Optional: set encryption options for older Excel formats
            // workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook to another memory stream
            MemoryStream encryptedStream = workbook.SaveToStream();
            long encryptedLength = encryptedStream.Length;
            Console.WriteLine($"Encrypted stream length: {encryptedLength} bytes");

            // Compare lengths to confirm encryption overhead
            long overhead = encryptedLength - unencryptedLength;
            Console.WriteLine($"Encryption overhead: {overhead} bytes");

            // Clean up streams
            unencryptedStream.Dispose();
            encryptedStream.Dispose();
        }
    }
}
