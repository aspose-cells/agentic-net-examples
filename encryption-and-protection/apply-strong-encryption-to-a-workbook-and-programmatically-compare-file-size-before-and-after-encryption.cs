// Title: C# Example: Strongly Encrypt an Aspose.Cells Workbook and Compare File Sizes
// Description: Shows how to create an Excel workbook with Aspose.Cells, save it unencrypted, apply a password and 128‑bit StrongCryptographicProvider encryption, save the protected file, measure the size before and after encryption, display the difference, and reload the workbook to confirm access—all in C#.
// Keywords: Aspose.Cells | C# encryption example | StrongCryptographicProvider | Workbook password protection | Excel file size comparison | SetEncryptionOptions | LoadOptions password | secure Excel file .NET | Aspose.Cells GitHub sample | Excel encryption API
// Common Searches: How to encrypt an Excel workbook with Aspose.Cells C# | Aspose.Cells StrongCryptographicProvider example | Measure file size before and after Excel encryption | Load password‑protected workbook using Aspose.Cells | C# code to apply 128‑bit encryption to .xlsx
// Developer Intent: Apply strong password‑based encryption to an Excel workbook with Aspose.Cells and evaluate the resulting file‑size change.
// Use Cases: Protect confidential spreadsheet data before distribution using 128‑bit encryption. | Quantify storage overhead introduced by workbook encryption for capacity planning. | Automate verification that an encrypted workbook can be opened programmatically with the correct password.
// AI Prompts: Generate C# code that encrypts an Aspose.Cells workbook with a 256‑bit AES key and reports the size delta. | Explain the parameters of SetEncryptionOptions in Aspose.Cells and how to validate the encrypted file. | Suggest a logging format that records encryption algorithm, key size, and file‑size comparison results.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create an Excel workbook with Aspose.Cells, save it unencrypted, apply a password and 128‑bit StrongCryptographicProvider encryption, save the protected file, measure the size before and after encryption, display the difference, and reload the workbook to confirm access—all in C#.
class Program
{
    static void Main()
    {
        // Paths for the unencrypted and encrypted workbooks
        string unencryptedPath = "UnencryptedWorkbook.xlsx";
        string encryptedPath = "EncryptedWorkbook.xlsx";

        // -------------------- Create Workbook --------------------
        Workbook workbook = new Workbook(); // create a new workbook

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Encryption test");

        // -------------------- Save Unencrypted Workbook --------------------
        workbook.Save(unencryptedPath); // save without any protection

        // Get file size before encryption
        long sizeBefore = new FileInfo(unencryptedPath).Length;

        // -------------------- Apply Strong Encryption --------------------
        // Set password for opening the workbook
        workbook.Settings.Password = "StrongPassword123";

        // Set encryption options: StrongCryptographicProvider with 128-bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // -------------------- Save Encrypted Workbook --------------------
        workbook.Save(encryptedPath); // save the encrypted version

        // Get file size after encryption
        long sizeAfter = new FileInfo(encryptedPath).Length;

        // -------------------- Compare File Sizes --------------------
        Console.WriteLine($"Size before encryption: {sizeBefore} bytes");
        Console.WriteLine($"Size after encryption:  {sizeAfter} bytes");
        Console.WriteLine($"Size increase: {sizeAfter - sizeBefore} bytes");

        // -------------------- Verify Encrypted Workbook Can Be Loaded --------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "StrongPassword123";

        Workbook loadedEncrypted = new Workbook(encryptedPath, loadOptions);
        Console.WriteLine($"Loaded encrypted workbook, cell A1 value: {loadedEncrypted.Worksheets[0].Cells["A1"].Value}");

        // Clean up
        workbook.Dispose();
        loadedEncrypted.Dispose();
    }
}
