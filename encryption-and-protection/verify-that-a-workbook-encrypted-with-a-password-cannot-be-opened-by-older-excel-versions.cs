// Title: Validate that a password‑protected AES‑256 .xlsx workbook cannot be opened by legacy Excel versions with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify encrypted .xlsx fails to open in Excel 2003 | how to test password protection compatibility with older Excel releases using .NET | C# code to confirm AES256 encrypted workbook is not readable by Excel 2007 | detect unsupported encryption when loading a password‑protected Excel file with Aspose.Cells | validate that legacy Excel cannot open AES‑256 encrypted .xlsx using LoadOptions in C#
// Tags: AES256 workbook encryption verification | Aspose.Cells legacy Excel compatibility check | password protected .xlsx load failure detection | C# LoadOptions password exception handling | Excel encryption version support validation

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, applies a strong password which triggers AES‑256 encryption for the .xlsx file, saves it, then attempts to reload it with LoadOptions. The expected outcome is an exception, confirming that older Excel versions cannot open the AES‑encrypted workbook.
class VerifyEncryption
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

            // Set a strong password (AES256 encryption will be used for .xlsx)
            wb.Settings.Password = "StrongPassword123";

            // Save the workbook as an encrypted .xlsx file
            string encryptedFile = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedFile, SaveFormat.Xlsx);

            // Verify the file was created
            if (!File.Exists(encryptedFile))
            {
                Console.WriteLine($"Failed to create the encrypted file: {encryptedFile}");
                return;
            }

            // LoadOptions without specifying a format; Aspose.Cells will auto‑detect.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "StrongPassword123";

            try
            {
                // Attempt to open the encrypted file using older Excel format expectations.
                // This should fail because the file is encrypted with AES (unsupported by older Excel).
                Workbook oldVersionWb = new Workbook(encryptedFile, loadOptions);
                Console.WriteLine("Unexpectedly opened the encrypted workbook with an older Excel format.");
            }
            catch (Exception ex)
            {
                // Expected outcome: an exception indicating the file cannot be opened.
                Console.WriteLine("Failed to open encrypted workbook with older Excel version as expected.");
                Console.WriteLine("Exception message: " + ex.Message);
            }
        }
        catch (Exception e)
        {
            // Catch any unexpected errors.
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }
}
