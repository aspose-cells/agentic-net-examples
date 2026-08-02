// Title: Validate that a password‑protected Aspose.Cells workbook cannot be opened without the password (C#)
// Description: The example creates an Excel workbook, applies Workbook.Settings.Password, saves it, confirms encryption via Workbook.Settings.IsEncrypted and FileFormatUtil.DetectFileFormat, attempts to load the file without a password (expecting an exception), then opens it with LoadOptions.Password and reads a cell value.
// Keywords: Aspose.Cells C# encryption | Workbook.Settings.Password | Workbook.Settings.IsEncrypted | FileFormatUtil DetectFileFormat IsEncrypted | LoadOptions password protected workbook | verify encrypted Excel file | catch exception opening encrypted workbook | read cell from password protected workbook
// Common Searches: Aspose.Cells check if workbook is encrypted | open password protected Excel file with Aspose.Cells | detect encrypted Excel file using FileFormatUtil | exception when opening encrypted workbook without password | C# Aspose.Cells load options password example
// Developer Intent: Demonstrate that an Excel file encrypted with a password cannot be opened without supplying the password, and show the correct way to open it using LoadOptions.
// Use Cases: Automated validation that confidential workbooks are properly encrypted before distribution. | Batch processing pipelines that need to flag or reject password‑protected Excel files lacking decryption support. | Testing third‑party viewers to ensure they reject encrypted workbooks, preserving data security.
// AI Prompts: Generate C# code with Aspose.Cells that verifies an encrypted workbook throws an exception when opened without a password. | Show how to use FileFormatUtil.DetectFileFormat to read the IsEncrypted flag of a saved Excel file. | Explain how to open a password‑protected workbook using LoadOptions and retrieve a specific cell value.

using System;
using Aspose.Cells;
using System.IO;

// The example creates an Excel workbook, applies Workbook.Settings.Password, saves it, confirms encryption via Workbook.Settings.IsEncrypted and FileFormatUtil.DetectFileFormat, attempts to load the file without a password (expecting an exception), then opens it with LoadOptions.Password and reads a cell value.
class WorkbookEncryptionValidation
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "mySecret";

        // Save the encrypted workbook
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath);

        // Verify that the workbook reports itself as encrypted
        Console.WriteLine("Workbook Settings.IsEncrypted: " + wb.Settings.IsEncrypted);

        // Use FileFormatUtil to detect encryption at file level
        FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(encryptedPath);
        Console.WriteLine("FileFormatInfo.IsEncrypted: " + fileInfo.IsEncrypted);

        // Simulate a third‑party viewer that does NOT provide a password
        try
        {
            // Attempt to open the encrypted file without a password
            Workbook wbWithoutPassword = new Workbook(encryptedPath);
            Console.WriteLine("Opened without password (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected failure because the file is encrypted
            Console.WriteLine("Failed to open without password as expected: " + ex.Message);
        }

        // Open the workbook with the correct password
        LoadOptions loadOptions = new LoadOptions { Password = "mySecret" };
        Workbook wbWithPassword = new Workbook(encryptedPath, loadOptions);
        Console.WriteLine("Opened with password, cell A1 value: " + wbWithPassword.Worksheets[0].Cells["A1"].StringValue);
    }
}
