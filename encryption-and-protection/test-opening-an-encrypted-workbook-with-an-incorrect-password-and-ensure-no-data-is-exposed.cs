// Title: Validate that an encrypted Excel workbook cannot be opened with an incorrect password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to test opening a password‑protected XLSX with a wrong password | C# catch CellsException when loading an encrypted workbook with an incorrect password | verify that sensitive data is not exposed after failed decryption using Aspose.Cells | load encrypted Excel file with LoadOptions wrong password Aspose.Cells .NET example
// Tags: Aspose.Cells load encrypted workbook with wrong password | CellsException handling for decryption failure | LoadOptions password mismatch scenario | temporary encrypted file cleanup in C# | validate password protection of Excel workbook

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, writes sensitive data, encrypts it with a password, saves it as an XLSX file, then attempts to reopen it using LoadOptions with an incorrect password. It catches the expected CellsException, logs the failure, and finally deletes the temporary encrypted file to ensure no data remains exposed.
class EncryptedWorkbookTest
{
    static void Main()
    {
        // Path for the temporary encrypted workbook
        string filePath = "encrypted.xlsx";

        // -------------------- Create and encrypt workbook --------------------
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");
        sheet.Cells["B2"].PutValue(12345);

        // Set password to encrypt the workbook
        wb.Settings.Password = "CorrectPassword";

        // Save the encrypted workbook to disk
        wb.Save(filePath, SaveFormat.Xlsx);

        // -------------------- Attempt to open with incorrect password --------------------
        try
        {
            // LoadOptions with an incorrect password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "WrongPassword"
            };

            // Attempt to load the encrypted workbook
            Workbook wbWrong = new Workbook(filePath, loadOptions);

            // If loading succeeds (which it shouldn't), attempt to read data
            string cellValue = wbWrong.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Unexpectedly opened workbook. Cell A1 value: " + cellValue);
        }
        catch (CellsException ex)
        {
            // Expected exception for incorrect password
            Console.WriteLine("Failed to open workbook with incorrect password: " + ex.Message);
        }
        finally
        {
            // Clean up the temporary file
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }
}
