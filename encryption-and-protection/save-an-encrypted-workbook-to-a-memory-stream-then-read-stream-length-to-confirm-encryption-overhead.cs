// Title: Encrypt an Aspose.Cells Workbook to a MemoryStream and Check Size (C#)
// Description: Creates a workbook, adds sample data, applies password protection with StrongCryptographicProvider (128‑bit), saves the encrypted file to a MemoryStream, reads its Length, then saves a plain version to another stream to illustrate the encryption overhead.
// Keywords: Aspose.Cells encrypt workbook C# | save encrypted Excel to MemoryStream | Workbook.Settings.Password | SetEncryptionOptions Aspose.Cells | measure encrypted stream size | encryption overhead Excel .NET
// Common Searches: how to encrypt Excel with Aspose.Cells and stream it | memory stream length of encrypted workbook Aspose | compare encrypted vs plain workbook size .NET | set encryption type and key length for .xls using Aspose
// Developer Intent: Save a password‑protected workbook to a stream and determine the byte increase caused by encryption.
// Use Cases: Validate the byte overhead introduced by Excel encryption before transmitting files. | Generate an encrypted workbook entirely in memory for HTTP responses or API returns. | Programmatically switch between encrypted and unencrypted outputs using the same Workbook instance.
// AI Prompts: Write C# code that encrypts an Aspose.Cells workbook with a custom password, uses StrongCryptographicProvider 256‑bit encryption, saves it to a MemoryStream, and returns the stream length. | Explain why an encrypted Excel stream is larger than an unencrypted one and how to calculate the overhead. | Show how to clear Workbook.Settings.Password after encryption and reuse the workbook to produce an unencrypted MemoryStream.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds sample data, applies password protection with StrongCryptographicProvider (128‑bit), saves the encrypted file to a MemoryStream, reads its Length, then saves a plain version to another stream to illustrate the encryption overhead.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["B1"].PutValue(42);

        // Apply password protection (encryption)
        workbook.Settings.Password = "securePwd";

        // Optional: specify encryption type and key length for .xls format
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to a memory stream
        MemoryStream encryptedStream = workbook.SaveToStream();

        // Output the length of the encrypted stream (shows encryption overhead)
        Console.WriteLine($"Encrypted stream length: {encryptedStream.Length} bytes");

        // For comparison, save an unencrypted version to another stream
        workbook.Settings.Password = null; // remove password
        MemoryStream plainStream = workbook.SaveToStream();
        Console.WriteLine($"Unencrypted stream length: {plainStream.Length} bytes");
    }
}
