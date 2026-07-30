// Title: Save a Password‑Protected Aspose.Cells Workbook to a MemoryStream and Measure Encryption Overhead (C#)
// Description: Demonstrates how to create an Aspose.Cells Workbook, write data, save it to an unencrypted MemoryStream, apply a password with StrongCryptographicProvider (128‑bit), save the encrypted version to another MemoryStream, compare the two stream lengths to reveal encryption overhead, check the Settings.IsEncrypted flag, and reload the protected file using LoadOptions.Password.
// Keywords: Aspose.Cells C# | Workbook.SaveToStream | password protection | encryption overhead | MemoryStream Excel | StrongCryptographicProvider | 128‑bit encryption | IsEncrypted flag | LoadOptions.Password | compare encrypted and unencrypted workbook size
// Common Searches: Aspose.Cells save encrypted workbook to MemoryStream | measure Excel encryption overhead with Aspose.Cells | load password‑protected workbook from stream C# | Workbook.IsEncrypted after SaveToStream | how to use StrongCryptographicProvider in Aspose.Cells
// Developer Intent: The developer needs to persist a workbook with password protection to a MemoryStream, verify that encryption adds extra bytes, and confirm that the file can be opened again using the supplied password.
// Use Cases: Transmit an Excel file over a network after encrypting it in memory, while tracking the size impact of encryption. | Automated tests that validate encryption settings by checking Settings.IsEncrypted before and after saving. | Server‑side processing that loads an encrypted workbook from a stream without writing to disk.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook with a password to a MemoryStream and prints the size difference between the encrypted and unencrypted streams. | Create a reusable method that accepts a Workbook and a password, returns an encrypted MemoryStream, and outputs the encryption overhead in bytes. | Explain step‑by‑step how to verify a workbook’s encrypted state using Settings.IsEncrypted after saving and after loading with LoadOptions.Password.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create an Aspose.Cells Workbook, write data, save it to an unencrypted MemoryStream, apply a password with StrongCryptographicProvider (128‑bit), save the encrypted version to another MemoryStream, compare the two stream lengths to reveal encryption overhead, check the Settings.IsEncrypted flag, and reload the protected file using LoadOptions.Password.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sample data");

        // Save the workbook without protection to a memory stream
        MemoryStream unprotectedStream = wb.SaveToStream();
        long unprotectedLength = unprotectedStream.Length;

        // Apply password protection (encryption)
        wb.Settings.Password = "mySecretPwd";
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the protected workbook to another memory stream
        MemoryStream protectedStream = wb.SaveToStream();
        long protectedLength = protectedStream.Length;

        // Output the lengths to demonstrate encryption overhead
        Console.WriteLine($"Unprotected stream length: {unprotectedLength}");
        Console.WriteLine($"Protected stream length:   {protectedLength}");
        Console.WriteLine($"Encryption overhead:       {protectedLength - unprotectedLength}");

        // Verify that the protected workbook reports as encrypted
        Console.WriteLine($"Workbook.IsEncrypted after protection: {wb.Settings.IsEncrypted}");

        // Load the protected workbook from the stream using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "mySecretPwd";
        Workbook loadedWorkbook = new Workbook(protectedStream, loadOptions);
        Console.WriteLine($"Loaded workbook IsEncrypted: {loadedWorkbook.Settings.IsEncrypted}");

        // Clean up streams
        unprotectedStream.Dispose();
        protectedStream.Dispose();
    }
}
