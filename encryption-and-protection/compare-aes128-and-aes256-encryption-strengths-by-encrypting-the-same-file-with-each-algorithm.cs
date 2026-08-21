// Title: Compare AES‑128 vs AES‑256 Encryption of Excel Workbooks with Aspose.Cells for .NET
// Description: Creates a sample workbook, then encrypts the same file twice—once with a 128‑bit AES key and once with a 256‑bit AES key—using Aspose.Cells' StrongCryptographicProvider. Each file is saved, reopened with the password to confirm encryption, and the paths are printed for easy comparison.
// Keywords: Aspose.Cells | C# | AES 128 encryption | AES 256 encryption | Excel workbook encryption | StrongCryptographicProvider | SetEncryptionOptions | Password‑protected Excel file | compare encryption strength | .NET encryption example
// Common Searches: encrypt Excel file with AES‑128 using Aspose.Cells | AES‑256 workbook protection Aspose.Cells .NET | difference between AES‑128 and AES‑256 in Excel | verify password‑protected workbook Aspose.Cells | set encryption key length Aspose.Cells C#
// Developer Intent: Encrypt the same Excel workbook with 128‑bit and 256‑bit AES keys using Aspose.Cells and validate that both files open with the password.
// Use Cases: Generate a moderately secured workbook (AES‑128) for internal distribution. | Create a highly secured workbook (AES‑256) for confidential data exchange. | Programmatically confirm that the selected encryption level was applied correctly.
// AI Prompts: Write C# code that encrypts an existing Excel file with AES‑256 using Aspose.Cells and saves it with a password. | Show how to compare the file sizes of AES‑128 and AES‑256 encrypted workbooks created with Aspose.Cells. | Provide a try‑catch example for handling exceptions when opening a password‑protected workbook encrypted with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionComparison
{
    // Creates a sample workbook, then encrypts the same file twice—once with a 128‑bit AES key and once with a 256‑bit AES key—using Aspose.Cells' StrongCryptographicProvider. Each file is saved, reopened with the password to confirm encryption, and the paths are printed for easy comparison.
    class Program
    {
        static void Main()
        {
            // Path for the original workbook (unencrypted)
            string originalPath = "OriginalWorkbook.xlsx";

            // Create a workbook and add sample data
            Workbook originalWorkbook = new Workbook();
            Worksheet sheet = originalWorkbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Strength Comparison");
            sheet.Cells["A2"].PutValue("AES-128 vs AES-256");
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Save the original workbook (optional, for reference)
            originalWorkbook.Save(originalPath, SaveFormat.Xlsx);

            // Password to protect the workbooks
            string password = "StrongPassword123";

            // ---------- AES-128 Encryption ----------
            // Load the original workbook to ensure identical content
            Workbook workbook128 = new Workbook(originalPath);
            // Set password required to open the workbook
            workbook128.Settings.Password = password;
            // Apply encryption options: StrongCryptographicProvider with 128‑bit key
            workbook128.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
            // Save the AES‑128 encrypted workbook
            string encrypted128Path = "Encrypted_AES128.xlsx";
            workbook128.Save(encrypted128Path, SaveFormat.Xlsx);

            // Verify AES‑128 encrypted file can be opened
            LoadOptions loadOptions128 = new LoadOptions { Password = password };
            Workbook loaded128 = new Workbook(encrypted128Path, loadOptions128);
            Console.WriteLine($"AES‑128 encrypted file loaded. IsEncrypted: {loaded128.Settings.IsEncrypted}");

            // ---------- AES-256 Encryption ----------
            // Load the original workbook again for identical content
            Workbook workbook256 = new Workbook(originalPath);
            // Set password required to open the workbook
            workbook256.Settings.Password = password;
            // Apply encryption options: StrongCryptographicProvider with 256‑bit key
            workbook256.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
            // Save the AES‑256 encrypted workbook
            string encrypted256Path = "Encrypted_AES256.xlsx";
            workbook256.Save(encrypted256Path, SaveFormat.Xlsx);

            // Verify AES‑256 encrypted file can be opened
            LoadOptions loadOptions256 = new LoadOptions { Password = password };
            Workbook loaded256 = new Workbook(encrypted256Path, loadOptions256);
            Console.WriteLine($"AES‑256 encrypted file loaded. IsEncrypted: {loaded256.Settings.IsEncrypted}");

            // Comparison output (both files contain the same data, but use different key lengths)
            Console.WriteLine("Encryption comparison completed:");
            Console.WriteLine($"- AES‑128 file: {encrypted128Path}");
            Console.WriteLine($"- AES‑256 file: {encrypted256Path}");
        }
    }
}
