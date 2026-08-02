// Title: Encrypt an Excel workbook with a Unicode (Chinese/emoji) password and verify decryption using Aspose.Cells for .NET
// Description: This C# sample creates a workbook, writes a value to cell A1, protects it with a password that includes Chinese characters and an emoji, optionally enables AES‑128 encryption, saves the file, checks the IsEncrypted flag, reloads the workbook with LoadOptions, and reads the cell to confirm successful decryption.
// Keywords: Aspose.Cells | C# | Unicode password encryption | Chinese characters password Excel | emoji password | AES 128 encryption | Workbook.IsEncrypted | LoadOptions password | Excel file protection | international password support
// Common Searches: Aspose.Cells protect Excel with Chinese password | Encrypt Excel workbook using emoji password .NET | Load encrypted workbook Unicode password Aspose | Check if Excel file is encrypted after save | Set AES 128 encryption for Aspose.Cells workbook
// Developer Intent: Protect an Excel file with a non‑ASCII password and programmatically confirm it can be opened.
// Use Cases: Compliance‑driven reports that require multilingual or symbolic passwords. | Automated pipelines that store Excel assets with Unicode passwords and need validation of encryption status. | Cross‑regional applications where users may choose emojis or native language characters for workbook protection.
// AI Prompts: Generate C# code that encrypts an Aspose.Cells workbook with a Chinese/emoji password, uses AES‑128, saves it, and opens it with LoadOptions. | Show how to verify Workbook.Settings.IsEncrypted before and after loading a password‑protected file. | Explain Aspose.Cells support for Unicode characters in workbook passwords and any known limitations.

using System;
using Aspose.Cells;

namespace AsposeCellsUnicodePasswordDemo
{
    // This C# sample creates a workbook, writes a value to cell A1, protects it with a password that includes Chinese characters and an emoji, optionally enables AES‑128 encryption, saves the file, checks the IsEncrypted flag, reloads the workbook with LoadOptions, and reads the cell to confirm successful decryption.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Unicode password test");

            // Set a password that contains Unicode characters (e.g., Chinese characters and an emoji)
            string unicodePassword = "密码🔒";
            wb.Settings.Password = unicodePassword;

            // Optionally set stronger encryption (AES 128-bit) – not required but demonstrates usage
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string filePath = "UnicodeEncryptedWorkbook.xlsx";
            wb.Save(filePath);

            // Verify that the workbook reports being encrypted
            Console.WriteLine($"IsEncrypted after save: {wb.Settings.IsEncrypted}");

            // Load the encrypted workbook using the same Unicode password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = unicodePassword;
            Workbook loadedWb = new Workbook(filePath, loadOptions);

            // Confirm that the workbook was successfully decrypted
            Console.WriteLine($"IsEncrypted after load: {loadedWb.Settings.IsEncrypted}");

            // Verify the cell value to ensure correct decryption
            string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Decrypted cell value: {cellValue}");
        }
    }
}
