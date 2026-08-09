// Title: Encrypt an Excel workbook with password and AES‑128 using Aspose.Cells for .NET and test third‑party loading
// Description: Creates a workbook, adds data, applies a strong password with AES‑128 encryption, saves the file, confirms encryption via FileFormatUtil, attempts to open it with LoadOptions (simulating a third‑party library), and verifies the password without fully loading the workbook.
// Keywords: Aspose.Cells encrypt Excel | C# password protected workbook | AES 128 Excel encryption .NET | verify encrypted Excel file | LoadOptions password protected workbook | FileFormatUtil verify password | third‑party Excel decryption test
// Common Searches: how to encrypt an .xlsx with Aspose.Cells C# | set AES 128 encryption for Excel using Aspose | check if Excel file is encrypted without opening it | load password protected workbook with LoadOptions Aspose | verify Excel password programmatically C#
// Developer Intent: Generate a password‑protected, AES‑128 encrypted Excel file and confirm that it can be opened with a password‑aware load operation, then validate the password without fully loading the workbook.
// Use Cases: Protect confidential reports before distribution while ensuring only authorized code can open them. | Automate pre‑flight checks that confirm encryption status and password validity. | Validate compatibility of encrypted workbooks with external or third‑party processing tools.
// AI Prompts: Write C# code with Aspose.Cells to encrypt an .xlsx workbook using AES‑128 and a password, then detect its encryption state. | Show how to verify a password for an encrypted Excel file without opening it, using Aspose.Cells FileFormatUtil. | Provide a sample that loads a password‑protected workbook via LoadOptions and handles incorrect‑password exceptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Creates a workbook, adds data, applies a strong password with AES‑128 encryption, saves the file, confirms encryption via FileFormatUtil, attempts to open it with LoadOptions (simulating a third‑party library), and verifies the password without fully loading the workbook.
    class Program
    {
        static void Main()
        {
            // ------------------------------
            // 1. Create a new workbook and add sample data
            // ------------------------------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // ------------------------------
            // 2. Apply password protection and encryption options
            // ------------------------------
            // Set the password that will be required to open the file
            workbook.Settings.Password = "StrongPassword123";

            // Optionally set stronger encryption (AES 128-bit) – the EncryptionType is ignored for .xlsx but kept for completeness
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // ------------------------------
            // 3. Save the encrypted workbook
            // ------------------------------
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath);                           // save

            // ------------------------------
            // 4. Verify that the file is indeed encrypted using Aspose utilities
            // ------------------------------
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(encryptedPath);
            Console.WriteLine($"IsEncrypted (FileFormatInfo): {info.IsEncrypted}");

            // ------------------------------
            // 5. Attempt to open the encrypted file with a "third‑party" approach
            //    Here we simulate a third‑party library by using Aspose's LoadOptions with the password.
            //    In a real scenario, replace this block with the actual third‑party API call.
            // ------------------------------
            try
            {
                // LoadOptions with the correct password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = "StrongPassword123"
                };

                // Attempt to load the encrypted workbook
                Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);
                Console.WriteLine("Third‑party load succeeded. Workbook opened successfully.");
                Console.WriteLine($"First cell value: {loadedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Third‑party load failed:");
                Console.WriteLine(ex.Message);
            }

            // ------------------------------
            // 6. Demonstrate password verification without opening the file
            // ------------------------------
            using (Stream stream = File.OpenRead(encryptedPath))
            {
                bool passwordValid = FileFormatUtil.VerifyPassword(stream, "StrongPassword123");
                Console.WriteLine($"Password verification result: {passwordValid}");
            }
        }
    }
}
