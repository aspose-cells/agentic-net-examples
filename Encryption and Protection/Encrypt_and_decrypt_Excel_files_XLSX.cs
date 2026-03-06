using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using System.IO;

class ExcelEncryptionDemo
{
    static void Main()
    {
        // -------------------- Create and encrypt workbook --------------------
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Information");

        // Set a password for the workbook
        workbook.Settings.Password = "StrongPassword123";

        // Define encryption options (AES 128-bit)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        string encryptedFile = "EncryptedWorkbook.xlsx";
        workbook.Save(encryptedFile, SaveFormat.Xlsx);

        // -------------------- Verify encryption --------------------
        // Detect if the saved file is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedFile);
        Console.WriteLine($"Is '{encryptedFile}' encrypted? {formatInfo.IsEncrypted}");

        // -------------------- Decrypt (load) workbook --------------------
        // Load the encrypted workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "StrongPassword123";
        Workbook loadedWorkbook = new Workbook(encryptedFile, loadOptions);

        // Read the previously stored value
        string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine($"Decrypted cell value: {cellValue}");

        // -------------------- Remove protection and save unencrypted --------------------
        // Clear the password to remove protection
        loadedWorkbook.Settings.Password = null;

        // Save the workbook without encryption
        string unencryptedFile = "UnencryptedWorkbook.xlsx";
        loadedWorkbook.Save(unencryptedFile, SaveFormat.Xlsx);

        // Verify that the new file is not encrypted
        FileFormatInfo unencryptedInfo = FileFormatUtil.DetectFileFormat(unencryptedFile);
        Console.WriteLine($"Is '{unencryptedFile}' encrypted? {unencryptedInfo.IsEncrypted}");
    }
}