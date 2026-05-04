using System;
using Aspose.Cells;

class StrongEncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and add some sensitive data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

        // Set a password required to open the workbook
        workbook.Settings.Password = "StrongPassword!123";

        // Apply strong encryption (AES 256) to the workbook
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

        // Save the encrypted workbook
        string encryptedFile = "StrongEncryptedWorkbook.xlsx";
        workbook.Save(encryptedFile, SaveFormat.Xlsx);

        // Attempt to open the workbook without a password (should fail)
        try
        {
            Workbook invalidLoad = new Workbook(encryptedFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Opening without password failed: " + ex.Message);
        }

        // Load the workbook with the correct password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "StrongPassword!123";
        Workbook loadedWorkbook = new Workbook(encryptedFile, loadOptions);

        // Verify that the data is accessible after providing the password
        Console.WriteLine("Loaded cell value: " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);
    }
}