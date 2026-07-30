// Title: Encrypt an Excel workbook with a Unicode (Chinese & emoji) password and validate decryption – Aspose.Cells for .NET
// Description: This C# sample shows how to protect an .xlsx file using a password that includes Chinese characters and an emoji. It applies a 128‑bit StrongCryptographicProvider encryption, saves the workbook, reloads it with matching LoadOptions, checks the IsEncrypted flag, and reads a cell to confirm successful decryption.
// Keywords: Aspose.Cells | C# | Unicode password | Chinese password | emoji password | Excel encryption | strong cryptographic provider | 128‑bit encryption | LoadOptions | Workbook.IsEncrypted | .xlsx protection | non‑ASCII password
// Common Searches: Aspose.Cells encrypt Excel file with Chinese characters password | How to use emoji in Excel workbook password with Aspose.Cells | Set 128‑bit encryption for .xlsx using Aspose.Cells C# | Load password‑protected workbook with Unicode password Aspose.Cells | Check if workbook is encrypted after saving Aspose.Cells
// Developer Intent: Create a password‑protected workbook using Unicode characters, save it, then reopen it with the same password to verify decryption.
// Use Cases: Secure financial or HR reports that must be opened only with a non‑ASCII passphrase. | Meet regional compliance by allowing passwords in native scripts such as Chinese. | Implement strong 128‑bit encryption for Excel 2007+ files while supporting emoji‑based passwords. | Programmatically confirm that a saved workbook is encrypted and that its contents are correctly restored.
// AI Prompts: Write C# code with Aspose.Cells to encrypt an .xlsx file using a password containing Chinese characters and an emoji, then load and verify the content. | Explain how Aspose.Cells processes Unicode passwords during workbook encryption and decryption, including how to read the IsEncrypted property. | Provide step‑by‑step instructions for applying StrongCryptographicProvider (128‑bit) encryption to a workbook protected by a Unicode password.

using System;
using Aspose.Cells;

namespace AsposeCellsUnicodePasswordDemo
{
    // This C# sample shows how to protect an .xlsx file using a password that includes Chinese characters and an emoji. It applies a 128‑bit StrongCryptographicProvider encryption, saves the workbook, reloads it with matching LoadOptions, checks the IsEncrypted flag, and reads a cell to confirm successful decryption.
    class Program
    {
        static void Main()
        {
            // Unicode password (contains Chinese characters and an emoji)
            string unicodePassword = "密码🔒";

            // ------------------- Create and encrypt workbook -------------------
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["B2"].PutValue("Unicode password test");

            // Set the workbook encryption password
            wb.Settings.Password = unicodePassword;

            // (Optional) Set stronger encryption options for Excel 2007+ files
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedFile = "UnicodeEncryptedWorkbook.xlsx";
            wb.Save(encryptedFile, SaveFormat.Xlsx);

            // ------------------- Load and verify decryption -------------------
            // Prepare load options with the same Unicode password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = unicodePassword;

            // Load the password‑protected workbook
            Workbook loadedWb = new Workbook(encryptedFile, loadOptions);

            // Verify that the workbook is indeed encrypted
            Console.WriteLine("IsEncrypted: " + loadedWb.Settings.IsEncrypted);

            // Verify the data inside the workbook
            string cellValue = loadedWb.Worksheets[0].Cells["B2"].StringValue;
            Console.WriteLine("Decrypted cell value: " + cellValue);
        }
    }
}
