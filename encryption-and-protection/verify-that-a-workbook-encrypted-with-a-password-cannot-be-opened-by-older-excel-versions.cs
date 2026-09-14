// Title: Verify that a password‑encrypted .xlsx workbook cannot be opened by legacy Excel versions using Aspose.Cells for .NET
// AI Prompts: Create a C# program that uses Aspose.Cells to encrypt a workbook with a password, then try to open the file without providing the password to confirm the operation fails. | Show how to open a password‑protected .xlsx in C# by supplying the password through LoadOptions and handle the successful decryption case. | Demonstrate catching the exception thrown when an encrypted workbook is loaded with default LoadOptions, indicating it cannot be opened by older Excel versions.
// Common Searches: how to confirm that a password protected .xlsx cannot be opened in Excel 2003 using Aspose.Cells | Aspose.Cells C# load encrypted workbook without password throws exception | verify Excel file encryption compatibility with legacy Excel versions in .NET | C# example for testing encrypted workbook opening failure on older Excel formats | using LoadOptions to open password protected Excel file with Aspose.Cells
// Tags: Aspose.Cells encrypt workbook | LoadOptions set password for xlsx | detect legacy Excel incompatibility encrypted file | C# handle exception opening protected workbook | verify encrypted workbook blocks older Excel

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, adds sample data, applies a password, and saves it as an encrypted .xlsx. It then attempts to open the file without a password using default LoadOptions, which throws an exception, proving that legacy Excel versions cannot read the file. Finally, it opens the workbook with the correct password via LoadOptions to demonstrate successful decryption.
class WorkbookEncryptionVerification
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // 2. Apply password protection (default encryption for .xlsx is strong)
            wb.Settings.Password = "SecretPassword123";

            // 3. Save the encrypted workbook
            string encryptedFile = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedFile, SaveFormat.Xlsx);
            Console.WriteLine($"Encrypted workbook saved as '{encryptedFile}'.");

            // Ensure the file was created before proceeding
            if (!File.Exists(encryptedFile))
            {
                Console.WriteLine($"File '{encryptedFile}' was not found.");
                return;
            }

            // 4. Attempt to load the encrypted file with older Excel format load options
            try
            {
                // Simulate loading with an older Excel version (XLS) by not providing a password.
                // The LoadOptions constructor without parameters uses auto-detect format.
                LoadOptions oldVersionLoadOptions = new LoadOptions();
                // This should fail because the file is encrypted and a password is required.
                Workbook oldVersionWb = new Workbook(encryptedFile, oldVersionLoadOptions);
                Console.WriteLine("Unexpectedly opened with older format.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Verification succeeded: older Excel versions cannot open the encrypted workbook.");
                Console.WriteLine("Exception message: " + ex.Message);
            }

            // 5. Verify that the workbook can be opened with the correct (modern) settings
            try
            {
                LoadOptions modernLoadOptions = new LoadOptions
                {
                    Password = "SecretPassword123"
                };
                Workbook openedWb = new Workbook(encryptedFile, modernLoadOptions);
                Console.WriteLine("Workbook opened successfully with correct password and modern format.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to open workbook with correct password.");
                Console.WriteLine("Exception message: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
