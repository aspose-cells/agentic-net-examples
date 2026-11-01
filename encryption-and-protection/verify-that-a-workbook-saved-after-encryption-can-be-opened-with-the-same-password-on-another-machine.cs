// Title: Encrypt an XLSX workbook with a password and confirm it opens on another machine using Aspose.Cells for .NET
// Common Searches: how to set a password for an Excel workbook using Aspose.Cells C# | load a password‑protected XLSX file with Aspose.Cells LoadOptions C# | verify encrypted workbook can be opened on a different computer with Aspose.Cells | Aspose.Cells encrypt workbook and read cell after decryption C# example
// Tags: Aspose.Cells workbook password encryption C# | LoadOptions.Password for encrypted XLSX | SaveFormat.Xlsx encrypted workbook | cross‑machine workbook decryption verification | Aspose.Cells Settings.Password example

using System;
using Aspose.Cells;

// The sample creates a new workbook, writes "Hello, Aspose!" to cell A1, applies Settings.Password to encrypt it, saves the file as EncryptedWorkbook.xlsx, then reloads the file on another machine using LoadOptions with the same password and reads the cell value to confirm successful decryption.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and add a value to cell A1
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose!");

        // Set a password to encrypt the workbook
        workbook.Settings.Password = "MySecurePassword";

        // Save the encrypted workbook to a file
        string filePath = "EncryptedWorkbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the encrypted workbook using the same password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = "MySecurePassword";
        Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

        // Verify that the workbook opened correctly by reading the value from A1
        string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Value in A1 after decryption: " + cellValue);
    }
}
