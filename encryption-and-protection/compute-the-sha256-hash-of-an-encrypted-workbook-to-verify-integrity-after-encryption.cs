// Title: Compute SHA-256 checksum of an Aspose.Cells workbook saved to a MemoryStream in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, saves it to a MemoryStream with OoxmlSaveOptions, and returns the SHA‑256 hash of the stream as a hex string. | Show how to read the bytes from a MemoryStream containing an Aspose.Cells XLSX file and generate a SHA‑256 hash using .NET's SHA256 class. | Demonstrate adapting the workbook‑hash example to first encrypt the workbook with a password and then compute the SHA‑256 hash of the encrypted file.
// Common Searches: C# how to calculate SHA-256 hash of an Excel file created with Aspose.Cells | Aspose.Cells compute checksum of workbook saved to MemoryStream .NET | verify integrity of encrypted XLSX using SHA256 in C# | generate SHA256 hash from Aspose.Cells OoxmlSaveOptions output stream
// Tags: Aspose.Cells SHA256 checksum for XLSX memory stream | C# compute workbook hash with System.Security.Cryptography | validate encrypted Excel file integrity using .NET | OoxmlSaveOptions stream saving example | convert SHA256 byte array to string in C#

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample creates an Aspose.Cells workbook, adds data, saves it to a MemoryStream using OoxmlSaveOptions for XLSX format, computes the SHA‑256 hash of the stream bytes, outputs the hash as a hexadecimal string, and finally writes the workbook to disk.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose!");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure save options (no encryption for compatibility)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the workbook to a memory stream
            using (MemoryStream encryptedStream = new MemoryStream())
            {
                workbook.Save(encryptedStream, saveOptions);
                encryptedStream.Position = 0; // Reset stream position for hashing

                // Compute SHA‑256 hash of the workbook bytes
                byte[] hashBytes;
                using (SHA256 sha256 = SHA256.Create())
                {
                    hashBytes = sha256.ComputeHash(encryptedStream);
                }

                // Convert hash to a readable hex string
                string hashString = BitConverter.ToString(hashBytes).Replace("-", string.Empty);
                Console.WriteLine($"SHA-256 hash of workbook: {hashString}");

                // Write the workbook to disk for verification
                string outputPath = "Workbook.xlsx";
                try
                {
                    File.WriteAllBytes(outputPath, encryptedStream.ToArray());
                    Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ioEx)
                {
                    Console.Error.WriteLine($"Failed to write file: {ioEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
