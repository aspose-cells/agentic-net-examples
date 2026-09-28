// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and verify its integrity with a SHA‑256 checksum
// AI Prompts: Set Workbook.Settings.Password to a secret, then save the workbook as Xlsx to create an encrypted file with Aspose.Cells. | Compute a SHA‑256 hash of the encrypted .xlsx file using .NET's SHA256 class and store the hex string for later verification. | Open a password‑protected workbook by supplying the password in LoadOptions, then read cell values to confirm successful decryption.
// Common Searches: how to password protect an Excel file with Aspose.Cells in C# | generate SHA256 checksum for an encrypted .xlsx file in .NET | verify integrity of a password‑protected workbook using Aspose.Cells | load a password‑protected Excel workbook with LoadOptions in C#
// Tags: Aspose.Cells workbook password encryption | C# generate cryptographic checksum for file | verify encrypted Excel file integrity | LoadOptions open password‑protected workbook | save workbook as encrypted Xlsx with Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The example creates a workbook, applies a password via Workbook.Settings.Password, saves it as an encrypted Xlsx file, computes a SHA‑256 checksum of the saved file, compares the checksum to confirm the file hasn't been altered, and demonstrates loading the protected workbook using LoadOptions with the same password.
class WorkbookEncryptionWithChecksum
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Set a password to encrypt the workbook
        // This will encrypt the file with the specified password when saved
        workbook.Settings.Password = "MySecretPassword";

        // Define the path for the encrypted file
        string encryptedFilePath = "EncryptedWorkbook.xlsx";

        // Save the workbook (it will be encrypted due to the password set above)
        workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

        // Compute a SHA256 checksum of the encrypted file
        string checksum = ComputeFileChecksum(encryptedFilePath);
        Console.WriteLine($"Checksum of encrypted file: {checksum}");

        // Example of verification: recompute checksum and compare
        string recomputedChecksum = ComputeFileChecksum(encryptedFilePath);
        if (checksum.Equals(recomputedChecksum, StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("Checksum verification passed. File integrity confirmed.");
        }
        else
        {
            Console.WriteLine("Checksum verification failed. File may have been tampered with.");
        }

        // Optional: demonstrate loading the encrypted workbook with the password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = "MySecretPassword"
        };
        Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);
        Console.WriteLine($"Loaded cell A1 value: {loadedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
    }

    // Helper method to compute SHA256 checksum of a file and return it as a hex string
    private static string ComputeFileChecksum(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(stream);
            StringBuilder sb = new StringBuilder(hash.Length * 2);
            foreach (byte b in hash)
                sb.AppendFormat("{0:x2}", b);
            return sb.ToString();
        }
    }
}
