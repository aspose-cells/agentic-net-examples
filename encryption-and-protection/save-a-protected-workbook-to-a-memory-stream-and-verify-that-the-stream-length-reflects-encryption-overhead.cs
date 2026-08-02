// Title: Save a Password‑Protected Aspose.Cells Workbook to a MemoryStream and Measure Encryption Overhead (C#)
// Description: Creates a workbook, writes sample data, saves it to an unprotected MemoryStream, applies a password with 128‑bit StrongCryptographicProvider encryption, saves the encrypted workbook to a second MemoryStream, prints both lengths, checks the IsEncrypted flag, and reloads the encrypted workbook from the stream using LoadOptions.
// Keywords: Aspose.Cells C# | save workbook to MemoryStream | encrypt Excel file Aspose.Cells | password protection StrongCryptographicProvider | encryption overhead size | Workbook.Settings.IsEncrypted | LoadOptions password | compare encrypted vs unencrypted stream length | in‑memory Excel encryption | Aspose.Cells encryption API
// Common Searches: Aspose.Cells save encrypted workbook to MemoryStream | How to measure encryption overhead in Aspose.Cells | Load password‑protected Excel from MemoryStream using Aspose.Cells | Check if Aspose.Cells workbook is encrypted after saving | C# example for workbook encryption with StrongCryptographicProvider
// Developer Intent: Save a workbook with password protection to a MemoryStream, confirm the encrypted stream is larger than the unencrypted one, and reload it using the correct password.
// Use Cases: Generate a secure Excel file entirely in memory for transmission over APIs or messaging systems. | Quantify the size impact of Aspose.Cells encryption for compliance or storage‑cost analysis. | Process an encrypted workbook without touching the file system by loading it directly from a MemoryStream.
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, encrypts it with a password, saves both encrypted and unencrypted versions to MemoryStream, and outputs the size difference. | Explain why the stream length increases when Aspose.Cells applies StrongCryptographicProvider encryption and how the overhead is calculated. | Generate a C# unit test that asserts the encrypted MemoryStream length is greater than the unencrypted length for the same workbook.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, writes sample data, saves it to an unprotected MemoryStream, applies a password with 128‑bit StrongCryptographicProvider encryption, saves the encrypted workbook to a second MemoryStream, prints both lengths, checks the IsEncrypted flag, and reloads the encrypted workbook from the stream using LoadOptions.
public class ProtectedWorkbookMemoryStreamDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B1"].PutValue(42);

            // Save the unprotected workbook to a memory stream
            MemoryStream unprotectedStream = workbook.SaveToStream();

            // Apply password protection and encryption options
            workbook.Settings.Password = "mySecretPassword";
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the protected (encrypted) workbook to another memory stream
            MemoryStream protectedStream = workbook.SaveToStream();

            // Output the lengths of both streams to demonstrate encryption overhead
            Console.WriteLine($"Unprotected stream length: {unprotectedStream.Length}");
            Console.WriteLine($"Protected stream length:   {protectedStream.Length}");
            Console.WriteLine($"Length difference (overhead): {protectedStream.Length - unprotectedStream.Length}");

            // Verify that the workbook reports being encrypted
            Console.WriteLine($"Workbook.Settings.IsEncrypted: {workbook.Settings.IsEncrypted}");

            // Reset stream position before loading
            protectedStream.Position = 0;

            // Load the encrypted workbook from the memory stream using the correct password
            LoadOptions loadOptions = new LoadOptions { Password = "mySecretPassword" };
            Workbook loadedWorkbook = new Workbook(protectedStream, loadOptions);

            // Confirm that the loaded workbook is also recognized as encrypted
            Console.WriteLine($"Loaded workbook.Settings.IsEncrypted: {loadedWorkbook.Settings.IsEncrypted}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ProtectedWorkbookMemoryStreamDemo.Run();
    }
}
