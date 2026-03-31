using System;
using Aspose.Cells;

class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add some data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive data");

        // Set a password – this encrypts the workbook using Excel's standard encryption
        wb.Settings.Password = "MySecretPwd";

        // Save the encrypted workbook
        string filePath = "encrypted_workbook.xlsx";
        wb.Save(filePath);

        // Load the encrypted workbook by providing the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MySecretPwd";
        Workbook wbLoaded = new Workbook(filePath, loadOptions);

        // Verify that the data can be read after decryption
        Console.WriteLine("Loaded cell value: " + wbLoaded.Worksheets[0].Cells["A1"].StringValue);
    }
}