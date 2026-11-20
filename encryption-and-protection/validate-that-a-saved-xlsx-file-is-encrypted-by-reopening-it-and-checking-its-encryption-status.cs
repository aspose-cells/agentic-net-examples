// Title: How to confirm an XLSX workbook is encrypted by reopening it with and without a password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify workbook encryption after saving | load encrypted Excel file without password using LoadOptions in .NET | check if saved XLSX is password protected with Aspose.Cells | detect encryption status of an Excel workbook programmatically in C# | validate password protection of a workbook using Aspose.Cells LoadOptions
// Tags: verify workbook encryption Aspose.Cells | load encrypted XLSX without password | password protection validation Aspose.Cells | check encryption status Excel .NET | LoadOptions encrypted workbook handling

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets a password via wb.Settings.Password to encrypt it, saves it as XLSX, then attempts to reload the file without a password (expecting an encryption exception) and finally reloads it with the correct password to confirm the workbook is properly encrypted and can be decrypted.
class EncryptionValidation
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "EncryptedWorkbook.xlsx";

        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add some data (optional)
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose!");

            // -----------------------------------------------------------------
            // 2. Set encryption options (password protection)
            // -----------------------------------------------------------------
            // Assign a password; this automatically encrypts the workbook when saved.
            wb.Settings.Password = "MySecretPassword";

            // -----------------------------------------------------------------
            // 3. Save the workbook (encrypted)
            // -----------------------------------------------------------------
            wb.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved and encrypted at '{filePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook creation/saving: {ex.Message}");
            return;
        }

        // -----------------------------------------------------------------
        // 4. Reopen the saved file without providing a password
        // -----------------------------------------------------------------
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load without password – structure is accessible, but content is protected.
            LoadOptions loadOptions = new LoadOptions(); // No password supplied
            Workbook loadedWb = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded without password (structure only).");
        }
        catch (CellsException)
        {
            Console.WriteLine("The workbook is encrypted and cannot be opened without a password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
        }

        // -----------------------------------------------------------------
        // 5. Validate that the workbook can be opened with the correct password
        // -----------------------------------------------------------------
        try
        {
            LoadOptions pwdLoadOptions = new LoadOptions { Password = "MySecretPassword" };
            Workbook pwdWb = new Workbook(filePath, pwdLoadOptions);
            Console.WriteLine("Workbook opened with password successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to open workbook with password: {ex.Message}");
        }
    }
}
