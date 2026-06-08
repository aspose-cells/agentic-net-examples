using System;
using Aspose.Cells;

namespace AsposeCellsPasswordChangeDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create and encrypt the workbook ----------
            Workbook wb = new Workbook();                         // create
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");        // add some data

            // Set initial password (encryption)
            wb.Settings.Password = "OldPassword";

            // Optional: specify encryption type and key length
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";
            wb.Save(filePath);                                   // save

            // ---------- Change the password without altering content ----------
            // Load the encrypted workbook using the old password
            LoadOptions loadOpts = new LoadOptions();
            loadOpts.Password = "OldPassword";
            Workbook wbLoaded = new Workbook(filePath, loadOpts); // load

            // Assign a new password; this re‑encrypts the file with the new password
            wbLoaded.Settings.Password = "NewPassword";

            // Save the workbook (overwrites the original file)
            wbLoaded.Save(filePath);                             // save

            // Verify that the new password works
            LoadOptions verifyOpts = new LoadOptions { Password = "NewPassword" };
            Workbook wbVerify = new Workbook(filePath, verifyOpts);
            Console.WriteLine("Cell A1 value after password change: " + wbVerify.Worksheets[0].Cells["A1"].StringValue);
        }
    }
}