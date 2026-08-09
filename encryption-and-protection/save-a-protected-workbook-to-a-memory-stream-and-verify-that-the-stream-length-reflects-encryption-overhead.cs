// Title: Encrypt an Aspose.Cells workbook, save to MemoryStream, and measure size overhead (C#)
// Description: Creates a workbook, writes sample data, saves it to a plain MemoryStream, applies a password with 128‑bit StrongCryptographicProvider encryption, saves the encrypted version to another MemoryStream, compares the two lengths, and confirms the encrypted file can be reopened with the correct password.
// Keywords: Aspose.Cells encryption C# | save workbook to MemoryStream | password protected Excel stream | encryption overhead measurement | StrongCryptographicProvider 128‑bit | load encrypted workbook from stream | Workbook.IsEncrypted check
// Common Searches: Aspose.Cells save encrypted workbook to MemoryStream | How to measure size increase after Excel encryption | C# load password protected workbook from stream | Set workbook password Aspose.Cells | Encryption overhead Aspose.Cells example
// Developer Intent: Encrypt a workbook, store it in memory, and verify that encryption adds bytes while ensuring the file can be opened with the supplied password.
// Use Cases: Benchmark storage impact of Excel encryption in in‑memory workflows | Validate that a password‑protected workbook can be reloaded from a stream | Compare plain vs encrypted stream sizes for performance tuning
// AI Prompts: Write C# code that encrypts an Aspose.Cells workbook with a 128‑bit StrongCryptographicProvider password, saves it to a MemoryStream, and prints the byte difference from the unencrypted stream. | Show how to load a password‑protected Aspose.Cells workbook from a MemoryStream and verify the IsEncrypted property. | Explain the steps to configure encryption options (type and key size) before saving an Aspose.Cells workbook to a stream.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, writes sample data, saves it to a plain MemoryStream, applies a password with 128‑bit StrongCryptographicProvider encryption, saves the encrypted version to another MemoryStream, compares the two lengths, and confirms the encrypted file can be reopened with the correct password.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Sample");
        ws.Cells["B1"].PutValue(42);

        // Save the workbook without encryption to a memory stream
        MemoryStream plainStream = wb.SaveToStream();
        long plainLength = plainStream.Length;

        // Apply password protection (encryption)
        wb.Settings.Password = "securePwd";
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to another memory stream
        MemoryStream encryptedStream = wb.SaveToStream();
        long encryptedLength = encryptedStream.Length;

        // Display lengths to show encryption overhead
        Console.WriteLine($"Unencrypted stream length: {plainLength}");
        Console.WriteLine($"Encrypted stream length:   {encryptedLength}");
        Console.WriteLine($"Encryption added {encryptedLength - plainLength} bytes.");

        // Verify that the encrypted stream can be loaded with the password
        LoadOptions loadOptions = new LoadOptions { Password = "securePwd" };
        Workbook loadedWb = new Workbook(encryptedStream, loadOptions);
        Console.WriteLine($"Loaded workbook IsEncrypted: {loadedWb.Settings.IsEncrypted}");
    }
}
