// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and confirm that loading it without the password raises a CellsException
// Common Searches: how to set a password on an Excel workbook with Aspose.Cells in C# | Aspose.Cells C# example for saving a workbook as encrypted XLSX | catch CellsException when opening a password‑protected Excel file without providing a password | verify that an encrypted XLSX cannot be opened without a password using Aspose.Cells | test decryption failure of a password‑protected workbook in .NET
// Tags: Workbook.Settings.Password encryption | save encrypted xlsx Aspose.Cells | load password protected workbook exception | CellsException handling for encrypted Excel | verify decryption failure .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, writes data, applies a password via Workbook.Settings.Password, saves it as an encrypted XLSX file, then attempts to load the file without supplying the password. The expected CellsException is caught, confirming that decryption fails when the password is omitted.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add sample data
            // -------------------------------------------------
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // -------------------------------------------------
            // 2. Encrypt the workbook with a password
            // -------------------------------------------------
            workbook.Settings.Password = "SecretPassword123";
            string encryptedPath = "EncryptedWorkbook.xlsx";

            // Save the encrypted workbook
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved and encrypted at: {encryptedPath}");

            // -------------------------------------------------
            // 3. Attempt to open the encrypted file without providing a password.
            //    Aspose.Cells will throw an exception because the workbook is protected.
            // -------------------------------------------------
            if (File.Exists(encryptedPath))
            {
                try
                {
                    // Attempt to load without password – expected to fail.
                    var protectedWorkbook = new Workbook(encryptedPath);
                    // If we reach this line, decryption unexpectedly succeeded.
                    Console.WriteLine("Unexpected success: Workbook opened without a password.");
                }
                catch (CellsException ex)
                {
                    // Expected path: decryption fails due to missing password.
                    Console.WriteLine("Decryption failed as expected when opening without a password.");
                    Console.WriteLine($"Exception message: {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Catch any other unexpected exceptions.
                    Console.WriteLine("An unexpected error occurred while opening the encrypted workbook.");
                    Console.WriteLine($"Exception message: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"File not found: {encryptedPath}");
            }
        }
        catch (Exception ex)
        {
            // General safety net for any unforeseen errors.
            Console.WriteLine("An error occurred during the demo execution.");
            Console.WriteLine($"Exception message: {ex.Message}");
        }
    }
}
