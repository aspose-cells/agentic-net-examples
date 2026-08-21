// Title: Encrypt an Excel workbook with a multilingual password and validate decryption – Aspose.Cells for .NET
// Description: Shows how to create a workbook, protect it with a password containing Chinese characters and an emoji, save the file, verify the encryption flag, then reopen it using the same credential to retrieve the original cell value.
// Keywords: Aspose.Cells | C# Excel encryption | non‑ASCII password | multilingual credential | IsEncrypted flag | load encrypted .xlsx | Excel security example | Aspose.Cells .NET
// Common Searches: Aspose.Cells set password with Chinese characters | C# encrypt Excel file using emoji | check workbook encryption status after save | open .xlsx protected by password with non‑Latin characters | example of multilingual password protection in Aspose.Cells
// Developer Intent: Secure a workbook with a password that includes non‑Latin characters and confirm programmatic access.
// Use Cases: Distribute confidential reports that require passwords in local languages or symbols. | Automated test suites that need to verify encryption metadata after saving. | International applications that read protected Excel files using culturally specific passwords.
// AI Prompts: Provide C# code to protect an Aspose.Cells workbook with a password containing Chinese characters and an emoji, then confirm the IsEncrypted property. | How can I load an encrypted .xlsx file in .NET when the password includes non‑ASCII symbols and read a specific cell? | What are the considerations and limitations when using multilingual passwords with Aspose.Cells workbook protection?

using System;
using Aspose.Cells;

namespace AsposeCellsUnicodePasswordDemo
{
    // Shows how to create a workbook, protect it with a password containing Chinese characters and an emoji, save the file, verify the encryption flag, then reopen it using the same credential to retrieve the original cell value.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Unicode password test");

            // Set a password that contains Unicode characters (e.g., Chinese characters and an emoji)
            string unicodePassword = "密码🔒";
            wb.Settings.Password = unicodePassword;

            // Save the encrypted workbook
            string encryptedFile = "UnicodeEncryptedWorkbook.xlsx";
            wb.Save(encryptedFile);

            // Verify that the workbook reports being encrypted
            Console.WriteLine($"IsEncrypted after save: {wb.Settings.IsEncrypted}");

            // Load the encrypted workbook using the same Unicode password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = unicodePassword;
            Workbook wbLoaded = new Workbook(encryptedFile, loadOptions);

            // Confirm that the workbook is recognized as encrypted after loading
            Console.WriteLine($"IsEncrypted after load: {wbLoaded.Settings.IsEncrypted}");

            // Verify that the data is correctly decrypted
            string cellValue = wbLoaded.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Decrypted cell value: {cellValue}");
        }
    }
}
