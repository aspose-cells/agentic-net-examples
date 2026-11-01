// Title: Change the password of an encrypted Excel workbook with Aspose.Cells for .NET without decrypting the file
// Common Searches: asp.net change password of encrypted xlsx using Aspose.Cells | update workbook encryption password programmatically in C# | how to re‑encrypt an Excel file with a new password using Aspose.Cells | load encrypted workbook with old password and save with new password Aspose.Cells | change Excel file password without opening its content C# Aspose
// Tags: Aspose.Cells workbook password change | encrypted xlsx re‑encryption .NET | LoadOptions password property Aspose.Cells | Workbook.Settings.Password update | change Excel encryption without decryption

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, encrypts it with an initial password, then reloads the encrypted file using LoadOptions with the old password, updates Workbook.Settings.Password to a new value, and saves the file again, effectively changing the workbook's password without decrypting its contents.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add some sample data
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // -----------------------------------------------------------------
            // 2. Encrypt the workbook with an initial password
            // -----------------------------------------------------------------
            wb.Settings.Password = "oldPassword";            // set encryption password
            string encryptedPath = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedPath);                           // save encrypted file

            // -----------------------------------------------------------------
            // 3. Change the password without decrypting the file content
            // -----------------------------------------------------------------
            if (!File.Exists(encryptedPath))
                throw new FileNotFoundException("Encrypted workbook not found.", encryptedPath);

            // Load the encrypted workbook using the current password
            LoadOptions loadOpts = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "oldPassword"
            };
            Workbook encryptedWb = new Workbook(encryptedPath, loadOpts);

            // Change the password by updating the workbook settings
            encryptedWb.Settings.Password = "newPassword";

            // Save the workbook; the file remains encrypted with the new password
            encryptedWb.Save(encryptedPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
