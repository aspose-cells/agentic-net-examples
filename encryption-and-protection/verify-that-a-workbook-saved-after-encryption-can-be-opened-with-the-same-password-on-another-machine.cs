// Title: How to encrypt an Excel workbook with a password using Aspose.Cells for .NET and verify it opens on another computer
// Common Searches: asp.net encrypt excel workbook with password using Aspose.Cells | load password protected xlsx on a different machine with Aspose.Cells C# | verify workbook encryption and decryption using LoadOptions in Aspose.Cells | c# example for saving encrypted excel file and reopening it with the same password | aspose.cells cross machine excel encryption validation
// Tags: Aspose.Cells password protection workbook | Save encrypted XLSX with Settings.Password | LoadOptions password decryption Aspose.Cells | Cross-machine Excel encryption verification | C# workbook encryption example

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, applying a password via Settings.Password, saving it as an encrypted XLSX, then loading it on a simulated different machine using LoadOptions with the same password and confirming the cell content matches the original.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        // Define file path and password
        string filePath = "EncryptedWorkbook.xlsx";
        string password = "SecurePass123";

        // -------------------------
        // Create a new workbook
        // -------------------------
        Workbook wb = new Workbook();
        // Write a test value to cell A1
        wb.Worksheets[0].Cells["A1"].PutValue("Hello, Aspose!");

        // Apply password protection (encryption)
        wb.Settings.Password = password;

        // Save the encrypted workbook to disk
        wb.Save(filePath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved and encrypted at: {filePath}");

        // -------------------------------------------------
        // Load the encrypted workbook on a different machine
        // -------------------------------------------------
        // Simulate a different environment by creating new LoadOptions
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            // Provide the same password used during encryption
            Password = password
        };

        // Load the workbook using the password
        Workbook loadedWb = new Workbook(filePath, loadOptions);
        Console.WriteLine("Workbook loaded successfully with the correct password.");

        // Verify that the content is accessible
        string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine($"Value in A1 after decryption: {cellValue}");

        // Simple validation
        if (cellValue == "Hello, Aspose!")
        {
            Console.WriteLine("Verification succeeded: the workbook opened with the correct password.");
        }
        else
        {
            Console.WriteLine("Verification failed: unexpected cell content.");
        }
    }
}
