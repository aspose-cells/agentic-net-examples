// Title: Verify that an Aspose.Cells workbook encrypted with a password cannot be opened using a similar but incorrect password in C#
// AI Prompts: Write C# code using Aspose.Cells to save a workbook with a password, then attempt to load it with a different password and capture the CellsException. | Show how to detect a failed decryption attempt by handling the exception thrown when an incorrect password is used with Aspose.Cells.
// Common Searches: Aspose.Cells C# open encrypted workbook with wrong password exception | catch CellsException when loading password protected Excel using Aspose.Cells | verify that similar password cannot decrypt Aspose.Cells workbook | test workbook encryption failure with wrong password in C#
// Tags: Aspose.Cells password protection validation | C# load encrypted workbook with incorrect password | CellsException handling for invalid password | Aspose.Cells workbook encryption test

using System;
using Aspose.Cells;

// The sample creates an Excel workbook, encrypts it with a password using Aspose.Cells, saves it, and then attempts to open the file with a similar but wrong password. The code catches the CellsException to confirm that the workbook cannot be decrypted with an incorrect password.
class WorkbookEncryptionTest
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "CorrectPassword123";
        string encryptedFile = "EncryptedWorkbook.xlsx";
        wb.Save(encryptedFile, SaveFormat.Xlsx);

        // Attempt to open the encrypted workbook with a different (similar) password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = "CorrectPassword124"; // similar but incorrect password

        try
        {
            // This should throw an exception because the password is wrong
            Workbook wbWrong = new Workbook(encryptedFile, loadOptions);
            Console.WriteLine("Workbook opened with wrong password (unexpected).");
        }
        catch (CellsException ex)
        {
            // Expected outcome: the workbook cannot be opened with the wrong password
            Console.WriteLine("Failed to open workbook with incorrect password: " + ex.Message);
        }
    }
}
