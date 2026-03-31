using System;
using Aspose.Cells;

class EncryptWorkbookExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Information");

        // Set a password that will be required to open the workbook
        workbook.Settings.Password = "MySecretPassword";

        // Specify encryption options (type is ignored for .xlsx but included for completeness)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the workbook as an encrypted XLSX file
        string outputPath = "EncryptedWorkbook.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Load the encrypted workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MySecretPassword";
        Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);

        // Verify that the data can be read after decryption
        Console.WriteLine("Cell A1 value: " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);
    }
}