// Title: Encrypt and Verify an Excel 97‑2003 Workbook with RC4 Using Aspose.Cells for .NET
// Description: Shows how to create a workbook, apply Office‑97/2000 compatible RC4 protection, save as .xls, reload with the password, and compare cell values to confirm data integrity.
// Keywords: Aspose.Cells | RC4 algorithm | Excel 97-2003 | C# .NET | password‑protected workbook | EncryptionType.Compatible | LoadOptions.Password | data integrity check
// Common Searches: RC4 encrypt Excel file C# Aspose.Cells | open password protected .xls with Aspose.Cells | verify data after decrypting Excel workbook | legacy Excel encryption .NET | set compatible encryption type Aspose
// Developer Intent: Apply RC4‑compatible protection, save the file, then open it with the password to ensure the original content remains unchanged.
// Use Cases: Secure legacy Excel 97‑2003 documents before distribution using a widely‑supported algorithm. | Automate validation that encrypted workbooks can be opened and retain exact data. | Batch‑process files that must meet older Office security requirements.
// AI Prompts: Generate C# code that encrypts an Excel workbook with RC4 via Aspose.Cells, saves it as .xls, and then reads it back to verify cell contents. | Explain the steps to set EncryptionType.Compatible and use LoadOptions.Password for a password‑protected workbook in Aspose.Cells. | Provide a method to confirm data consistency after decrypting an RC4‑encrypted Excel file using .NET.

using System;
using Aspose.Cells;

// Shows how to create a workbook, apply Office‑97/2000 compatible RC4 protection, save as .xls, reload with the password, and compare cell values to confirm data integrity.
class RC4EncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello RC4");

        // Set a password for the workbook
        string password = "rc4pwd";
        workbook.Settings.Password = password;

        // Apply RC4-compatible encryption (Office 97/2000 compatible)
        // EncryptionType.Compatible uses the older RC4 algorithm
        workbook.SetEncryptionOptions(EncryptionType.Compatible, 128);

        // Save the encrypted workbook (Excel 97-2003 format)
        string encryptedPath = "RC4EncryptedWorkbook.xls";
        workbook.Save(encryptedPath, SaveFormat.Excel97To2003);

        // Load the encrypted workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);

        // Retrieve the value to verify data consistency
        string originalValue = "Hello RC4";
        string loadedValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;

        // Output verification result
        Console.WriteLine("Original value: " + originalValue);
        Console.WriteLine("Loaded value:   " + loadedValue);
        Console.WriteLine("Data consistent: " + (originalValue == loadedValue));
    }
}
