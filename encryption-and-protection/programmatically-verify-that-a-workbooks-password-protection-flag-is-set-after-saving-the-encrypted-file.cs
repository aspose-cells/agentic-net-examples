using System;
using Aspose.Cells;

class VerifyWorkbookEncryption
{
    static void Main()
    {
        // Path for the encrypted workbook
        string filePath = "encryptedWorkbook.xlsx";

        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted data");

        // Set the password to encrypt the workbook
        workbook.Settings.Password = "mySecret";

        // Save the encrypted workbook
        workbook.Save(filePath);

        // Verify encryption flag using FileFormatInfo
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"FileFormatInfo.IsEncrypted: {formatInfo.IsEncrypted}");

        // Load the workbook with the correct password
        LoadOptions loadOptions = new LoadOptions { Password = "mySecret" };
        Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

        // Verify that the workbook reports it is encrypted
        Console.WriteLine($"Loaded workbook Settings.IsEncrypted: {loadedWorkbook.Settings.IsEncrypted}");
    }
}