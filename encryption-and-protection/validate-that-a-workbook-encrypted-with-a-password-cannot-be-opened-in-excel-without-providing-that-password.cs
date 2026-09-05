// Title: Verify that a password‑protected Excel workbook cannot be opened without the password using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program that assigns Settings.Password on a Workbook, saves it as an encrypted .xlsx file, then attempts to load the file without a password and captures the expected exception. | Write C# code that uses LoadOptions with the Password property to open the same encrypted workbook and reads a cell value to confirm successful decryption.
// Common Searches: C# Aspose.Cells test opening encrypted Excel file without providing password | how to catch exception when loading password protected workbook in Aspose.Cells | Aspose.Cells verify workbook encryption by trying to open without password | load password protected XLSX with LoadOptions password Aspose.Cells C# example | check that Settings.Password encrypts Excel file in Aspose.Cells .NET
// Tags: Aspose.Cells encrypt workbook password C# | Aspose.Cells open encrypted XLSX without password | Aspose.Cells LoadOptions password property | Aspose.Cells exception missing workbook password | C# verify Excel workbook encryption Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, sets Settings.Password to encrypt it, saves it as an .xlsx file, then tries to open the file without providing a password (capturing the thrown exception). It finally opens the same file with LoadOptions.Password and reads a cell to demonstrate successful decryption.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");
        sheet.Cells["A2"].PutValue(12345);

        // Set a password to encrypt the workbook
        wb.Settings.Password = "Secret123";

        // Save the encrypted workbook
        string encryptedPath = "EncryptedWorkbook.xlsx";
        wb.Save(encryptedPath, SaveFormat.Xlsx);

        // Attempt to open the encrypted workbook without providing a password
        try
        {
            // This should throw an exception because the password is required
            Workbook wbWithoutPassword = new Workbook(encryptedPath);
            Console.WriteLine("Opened without password (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to open without password as expected: " + ex.Message);
        }

        // Open the encrypted workbook with the correct password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = "Secret123"
        };
        Workbook wbWithPassword = new Workbook(encryptedPath, loadOptions);
        Console.WriteLine("Opened with password successfully.");
        // Verify that the data is accessible
        string cellValue = wbWithPassword.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Cell A1 value: " + cellValue);
    }
}
