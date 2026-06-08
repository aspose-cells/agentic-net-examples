using System;
using Aspose.Cells;

class VerifyEncryptionOlderVersion
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "Secret123";

        // Save the encrypted workbook in the modern .xlsx format
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath, SaveFormat.Xlsx);

        // Verify that the workbook reports as encrypted when opened with the correct password
        Workbook openedWithPassword = new Workbook(encryptedPath, new LoadOptions { Password = "Secret123" });
        Console.WriteLine("IsEncrypted (with correct password): " + openedWithPassword.Settings.IsEncrypted);

        // Attempt to open the encrypted file using an older Excel format (XLS) without providing a password
        try
        {
            LoadOptions oldFormatOptions = new LoadOptions(LoadFormat.Excel97To2003); // older Excel version
            Workbook oldVersionLoad = new Workbook(encryptedPath, oldFormatOptions);
            Console.WriteLine("Unexpectedly opened encrypted file with older Excel version.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to open encrypted workbook with older Excel version (expected): " + ex.Message);
        }

        // Open the encrypted workbook with the correct password using auto-detection
        LoadOptions correctPasswordOptions = new LoadOptions();
        correctPasswordOptions.Password = "Secret123";
        Workbook successfullyOpened = new Workbook(encryptedPath, correctPasswordOptions);
        Console.WriteLine("Successfully opened encrypted workbook. Cell A1 value: " + successfullyOpened.Worksheets[0].Cells["A1"].StringValue);
    }
}