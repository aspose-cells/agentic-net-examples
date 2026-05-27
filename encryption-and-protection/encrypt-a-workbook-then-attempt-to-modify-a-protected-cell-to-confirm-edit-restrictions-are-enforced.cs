using System;
using Aspose.Cells;

class WorkbookEncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and add initial data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Original Value");

        // Protect the worksheet with a password
        ws.Protect(ProtectionType.All, "sheetPwd", null);

        // Encrypt the workbook with a password
        wb.Settings.Password = "filePwd";
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted and protected workbook
        string filePath = "EncryptedProtectedWorkbook.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);
        wb.Dispose();

        // Load the workbook using the encryption password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "filePwd";
        Workbook loadedWb = new Workbook(filePath, loadOptions);
        Worksheet loadedWs = loadedWb.Worksheets[0];

        // Verify encryption and worksheet protection status
        Console.WriteLine("Workbook IsEncrypted: " + loadedWb.Settings.IsEncrypted);
        Console.WriteLine("Worksheet IsProtected: " + loadedWs.IsProtected);

        // Attempt to modify a protected cell without unprotecting
        string before = loadedWs.Cells["A1"].Value?.ToString();
        loadedWs.Cells["A1"].PutValue("Attempted Change");
        string after = loadedWs.Cells["A1"].Value?.ToString();

        // Display the result of the modification attempt
        Console.WriteLine($"Cell A1 before attempt: {before}");
        Console.WriteLine($"Cell A1 after attempt: {after}");
        Console.WriteLine("If the values are identical, the protection prevented the edit.");

        // Save the workbook (no changes will be persisted if protection blocked the edit)
        loadedWb.Save("ResultWorkbook.xlsx", SaveFormat.Xlsx);
        loadedWb.Dispose();
    }
}