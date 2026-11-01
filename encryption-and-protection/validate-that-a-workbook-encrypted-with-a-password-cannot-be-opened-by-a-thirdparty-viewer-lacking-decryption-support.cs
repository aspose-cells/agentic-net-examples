// Title: Verify that a password‑protected Excel workbook throws an exception when opened without a password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# test opening password‑protected Excel file without providing a password | Check exception thrown when loading an encrypted .xlsx with Aspose.Cells and no password | Validate workbook encryption by attempting to open an encrypted Excel file without a password in .NET | How to confirm that Aspose.Cells blocks unauthorized access to a password‑protected workbook | C# code to ensure an encrypted Excel workbook cannot be opened by a third‑party viewer
// Tags: Aspose.Cells workbook password encryption | exception when loading encrypted xlsx without password | LoadOptions password decryption Aspose.Cells | verify Excel file protection .NET | prevent unauthorized workbook access Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, applies a password via Workbook.Settings.Password, saves it, then attempts to load the file without a password, catching the expected exception, and finally opens the workbook correctly using LoadOptions with the password to confirm successful decryption.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Secret Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "MySecretPassword";

        // Save the encrypted workbook
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath);

        // Attempt to open the encrypted workbook without providing a password
        try
        {
            // This should throw an exception because the file is password‑protected
            Workbook wbWithoutPwd = new Workbook(encryptedPath);
            Console.WriteLine("Workbook opened without password (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected outcome: decryption fails
            Console.WriteLine("Failed to open encrypted workbook without password: " + ex.Message);
        }

        // Open the workbook with the correct password to verify successful decryption
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = "MySecretPassword" };
        Workbook wbWithPwd = new Workbook(encryptedPath, loadOptions);
        Console.WriteLine("Workbook opened with correct password. Cell A1 value: " + wbWithPwd.Worksheets[0].Cells["A1"].StringValue);
    }
}
