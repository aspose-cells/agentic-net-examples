using System;
using Aspose.Cells;

class WorkbookEncryptionValidation
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "mySecret";

        // Save the encrypted workbook
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath);

        // Verify that the workbook reports being encrypted
        Console.WriteLine($"IsEncrypted after save: {wb.Settings.IsEncrypted}");

        // Attempt to open the encrypted file without providing a password
        try
        {
            Workbook wbWithoutPwd = new Workbook(encryptedPath);
            // If no exception, check encryption flag (should be true)
            Console.WriteLine($"Loaded without password - IsEncrypted: {wbWithoutPwd.Settings.IsEncrypted}");
        }
        catch (Exception ex)
        {
            // Expected failure: third‑party viewer without decryption support cannot open it
            Console.WriteLine($"Failed to open without password: {ex.Message}");
        }

        // Open the encrypted workbook with the correct password using LoadOptions
        LoadOptions loadOptions = new LoadOptions { Password = "mySecret" };
        Workbook wbWithPwd = new Workbook(encryptedPath, loadOptions);
        Console.WriteLine($"Opened with password successfully. Cell A1 value: {wbWithPwd.Worksheets[0].Cells["A1"].StringValue}");
    }
}