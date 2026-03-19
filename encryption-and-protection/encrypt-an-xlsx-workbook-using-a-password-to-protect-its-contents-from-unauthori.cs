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

        // Set a password to encrypt the workbook
        wb.Settings.Password = "Secret123";

        // Optionally specify stronger encryption (128‑bit AES)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        string outputPath = "EncryptedWorkbook.xlsx";
        wb.Save(outputPath, SaveFormat.Xlsx);

        // Load the encrypted workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "Secret123";
        Workbook loadedWb = new Workbook(outputPath, loadOptions);

        // Verify that the data can be read after decryption
        string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Loaded cell value: " + cellValue);
    }
}