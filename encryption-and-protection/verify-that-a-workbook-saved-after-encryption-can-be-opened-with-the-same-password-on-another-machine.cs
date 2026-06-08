using System;
using Aspose.Cells;
using System.IO;

class VerifyEncryptedWorkbook
{
    static void Main()
    {
        // Path for the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";

        // ---------- Create and encrypt the workbook ----------
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Encrypted content");

        // Set the password to encrypt the workbook
        wb.Settings.Password = "MySecretPassword";

        // (Optional) Set stronger encryption options
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        wb.Save(filePath, SaveFormat.Xlsx);

        // ---------- Load the workbook on another machine ----------
        // Simulate loading on a different machine by using LoadOptions with the password
        LoadOptions loadOptions = new LoadOptions
        {
            Password = "MySecretPassword"
        };

        // Load the encrypted workbook
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // Verify that the workbook is indeed encrypted
        Console.WriteLine("IsEncrypted (original): " + wb.Settings.IsEncrypted);
        Console.WriteLine("IsEncrypted (loaded): " + loadedWb.Settings.IsEncrypted);

        // Verify that the data can be read correctly
        string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Cell A1 value after loading: " + cellValue);
    }
}