// Title: Compute SHA256 checksum of an encrypted Aspose.Cells workbook before and after decryption in C#
// Common Searches: how to calculate SHA256 hash of a password‑protected XLSX file using Aspose.Cells C# | verify workbook integrity after decrypting an encrypted Excel workbook with Aspose.Cells | C# example for loading an encrypted workbook with LoadOptions password and generating checksum | compare encrypted and decrypted workbook hashes using Aspose.Cells | remove password from Aspose.Cells workbook and recompute file hash
// Tags: calculate SHA256 hash of encrypted XLSX with Aspose.Cells | load encrypted workbook with password using LoadOptions | clear workbook password Aspose.Cells C# | checksum comparison of encrypted vs decrypted workbook | ensure Excel file integrity after decryption Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The example creates a workbook, encrypts it with a password, computes a SHA256 checksum of the encrypted byte stream, loads the workbook using LoadOptions with the password, removes the password to decrypt, saves the decrypted workbook, computes a second SHA256 checksum, and compares the two checksums to demonstrate how encryption changes the file content while allowing integrity verification.
class WorkbookChecksumDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B2"].PutValue(12345);
        sheet.Cells["C3"].PutValue(DateTime.Now);

        // Define password for encryption
        string password = "SecretPwd";

        // Save the workbook to a memory stream with encryption
        using (MemoryStream encryptedStream = new MemoryStream())
        {
            // Set the password in the workbook's settings
            wb.Settings.Password = password;
            wb.Save(encryptedStream, SaveFormat.Xlsx);
            byte[] encryptedBytes = encryptedStream.ToArray();

            // Compute checksum of the encrypted workbook
            string encryptedChecksum = ComputeSha256(encryptedBytes);
            Console.WriteLine($"Encrypted workbook checksum: {encryptedChecksum}");

            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            using (MemoryStream loadStream = new MemoryStream(encryptedBytes))
            {
                Workbook loadedWb = new Workbook(loadStream, loadOptions);

                // Remove the password to decrypt the workbook
                loadedWb.Settings.Password = null;

                // Save the decrypted workbook to another memory stream
                using (MemoryStream decryptedStream = new MemoryStream())
                {
                    loadedWb.Save(decryptedStream, SaveFormat.Xlsx);
                    byte[] decryptedBytes = decryptedStream.ToArray();

                    // Compute checksum of the decrypted workbook
                    string decryptedChecksum = ComputeSha256(decryptedBytes);
                    Console.WriteLine($"Decrypted workbook checksum: {decryptedChecksum}");

                    // Verify data consistency
                    if (encryptedChecksum != decryptedChecksum)
                    {
                        Console.WriteLine("Checksums differ as expected (encryption changes file content).");
                    }
                    else
                    {
                        Console.WriteLine("Checksums are identical (unexpected).");
                    }
                }
            }
        }
    }

    // Helper method to compute SHA256 checksum and return as hex string
    private static string ComputeSha256(byte[] data)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(data);
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hash)
                sb.Append(b.ToString("x2"));
            return sb.ToString();
        }
    }
}
