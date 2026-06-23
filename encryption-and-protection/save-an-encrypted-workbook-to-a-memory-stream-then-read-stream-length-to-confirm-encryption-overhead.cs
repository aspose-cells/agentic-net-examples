using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ---------- Create a plain (unencrypted) workbook ----------
        Workbook plainWb = new Workbook();
        Worksheet plainWs = plainWb.Worksheets[0];
        plainWs.Cells["A1"].PutValue("Hello");
        plainWs.Cells["A2"].PutValue("World");

        // Save the plain workbook to a memory stream
        MemoryStream plainStream = plainWb.SaveToStream();

        // ---------- Create an encrypted workbook ----------
        Workbook encryptedWb = new Workbook();
        Worksheet encryptedWs = encryptedWb.Worksheets[0];
        encryptedWs.Cells["A1"].PutValue("Hello");
        encryptedWs.Cells["A2"].PutValue("World");

        // Set a password to protect the workbook
        encryptedWb.Settings.Password = "myPassword";

        // Optionally specify encryption algorithm and key length
        encryptedWb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to a memory stream
        MemoryStream encryptedStream = encryptedWb.SaveToStream();

        // ---------- Compare stream lengths ----------
        Console.WriteLine($"Plain stream length: {plainStream.Length} bytes");
        Console.WriteLine($"Encrypted stream length: {encryptedStream.Length} bytes");
        Console.WriteLine($"Encryption overhead: {encryptedStream.Length - plainStream.Length} bytes");

        // Clean up
        plainStream.Dispose();
        encryptedStream.Dispose();
        plainWb.Dispose();
        encryptedWb.Dispose();
    }
}