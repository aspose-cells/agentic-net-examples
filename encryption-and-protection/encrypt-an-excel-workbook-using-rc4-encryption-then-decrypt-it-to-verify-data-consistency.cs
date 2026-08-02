// Title: Encrypt and Decrypt an Excel 97‑2003 Workbook with RC4 (XOR) using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply RC4‑style XOR encryption (128‑bit) with a password, save it as an XLS file, reload it using the same password, and verify that the cell data remains unchanged.
// Keywords: Aspose.Cells | C# | .NET | RC4 encryption | XOR encryption | Excel 97-2003 password protection | SetEncryptionOptions | LoadOptions password | data integrity verification | encrypted XLS file
// Common Searches: Aspose.Cells RC4 XOR encryption example | How to password‑protect an XLS file in C# | Load password protected Excel 97‑2003 workbook with Aspose.Cells | Verify data after decrypting an encrypted workbook | SetEncryptionOptions XOR 128‑bit Aspose.Cells
// Developer Intent: Apply RC4‑style XOR encryption to a legacy XLS workbook, save it, reopen it with the password, and confirm that the original content is preserved.
// Use Cases: Secure legacy Excel 97‑2003 files before distribution using password‑based RC4 encryption. | Automate validation of password‑protected workbooks in batch processing pipelines. | Ensure data integrity when encrypting and decrypting Excel files in .NET applications.
// AI Prompts: Generate C# code with Aspose.Cells that encrypts an XLS workbook using RC4 (XOR) encryption, saves it, then opens it with the password and checks a cell value for consistency. | Explain the role of SetEncryptionOptions and LoadOptions when working with RC4/XOR encryption in Aspose.Cells. | Provide a step‑by‑step tutorial for encrypting, saving, loading, and verifying an Excel 97‑2003 file with Aspose.Cells, including proper resource disposal.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Demonstrates how to create a workbook, apply RC4‑style XOR encryption (128‑bit) with a password, save it as an XLS file, reload it using the same password, and verify that the cell data remains unchanged.
    class Program
    {
        static void Main()
        {
            // Step 1: Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            const string originalValue = "RC4 Encryption Test";
            sheet.Cells["A1"].PutValue(originalValue);

            // Step 2: Set a password for the workbook
            // The password is required to open the encrypted file
            workbook.Settings.Password = "rc4pwd";

            // Step 3: Apply encryption options.
            // Excel 97‑2003 uses RC4 (XOR) encryption. 
            // Using EncryptionType.XOR with a key length of 128 bits mimics RC4 encryption.
            workbook.SetEncryptionOptions(EncryptionType.XOR, 128);

            // Step 4: Save the encrypted workbook in the older XLS format
            const string encryptedFilePath = "encrypted_rc4.xls";
            workbook.Save(encryptedFilePath, SaveFormat.Excel97To2003);

            // Step 5: Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "rc4pwd";
            Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);

            // Step 6: Verify that the data is consistent after decryption
            string loadedValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            bool isDataConsistent = string.Equals(originalValue, loadedValue, StringComparison.Ordinal);

            Console.WriteLine($"Original Value : {originalValue}");
            Console.WriteLine($"Loaded Value   : {loadedValue}");
            Console.WriteLine($"Data Consistency Check: {(isDataConsistent ? "PASS" : "FAIL")}");

            // Cleanup
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}
