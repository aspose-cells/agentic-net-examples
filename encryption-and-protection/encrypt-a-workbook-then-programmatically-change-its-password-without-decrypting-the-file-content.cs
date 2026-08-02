// Title: Replace an encrypted Excel workbook password using Aspose.Cells for .NET (no data decryption)
// Description: The C# sample creates a workbook, secures it with an initial password, reloads the file using that password, sets a new password via Settings.Password, and saves the file. Aspose.Cells automatically re‑encrypts the workbook with the new password while leaving all cell data unchanged.
// Keywords: Aspose.Cells password change | C# update Excel encryption | change workbook password without decrypting | re‑encrypt Excel file .NET | modify Excel protection Aspose | Excel password rotation | LoadOptions password Aspose.Cells
// Common Searches: how to change password of encrypted Excel file using Aspose.Cells C# | replace workbook encryption password without opening file | Aspose.Cells update Excel file password programmatically | change Excel workbook password after saving | C# re‑encrypt Excel workbook with new password
// Developer Intent: Swap the existing encryption password of an Excel workbook for a new one while keeping the file encrypted and without reading its contents.
// Use Cases: Rotate workbook passwords to satisfy compliance policies without re‑processing large data sets. | Update the protection key after a user changes credentials, preserving file confidentiality. | Automate bulk password resets for archived reports before moving them to long‑term storage.
// AI Prompts: Generate C# code that opens an encrypted Excel file with Aspose.Cells, changes its password, and saves it without accessing any cell values. | Explain how Aspose.Cells re‑encrypts a workbook when Settings.Password is modified. | Create a reusable method that takes a file path, old password, and new password and updates the workbook's encryption using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordChangeDemo
{
    // The C# sample creates a workbook, secures it with an initial password, reloads the file using that password, sets a new password via Settings.Password, and saves the file. Aspose.Cells automatically re‑encrypts the workbook with the new password while leaving all cell data unchanged.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create a new workbook and add some sample data
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();                         // create
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // -----------------------------------------------------------------
            // Step 2: Encrypt the workbook with the initial password
            // -----------------------------------------------------------------
            wb.Settings.Password = "OldPassword";                // set encryption password
            wb.Save("EncryptedWorkbook.xlsx");                    // save

            // -----------------------------------------------------------------
            // Step 3: Load the encrypted workbook using the old password
            // -----------------------------------------------------------------
            LoadOptions loadOpts = new LoadOptions();            // create load options
            loadOpts.Password = "OldPassword";                   // provide password to open
            Workbook loadedWb = new Workbook("EncryptedWorkbook.xlsx", loadOpts); // load

            // -----------------------------------------------------------------
            // Step 4: Change the password without decrypting the content
            // -----------------------------------------------------------------
            loadedWb.Settings.Password = "NewPassword";          // assign new password

            // -----------------------------------------------------------------
            // Step 5: Save the workbook; it will be re‑encrypted with the new password
            // -----------------------------------------------------------------
            loadedWb.Save("EncryptedWorkbook.xlsx");              // save (overwrite)

            // -----------------------------------------------------------------
            // Verification (optional): open with the new password
            // -----------------------------------------------------------------
            LoadOptions verifyOpts = new LoadOptions { Password = "NewPassword" };
            Workbook verifyWb = new Workbook("EncryptedWorkbook.xlsx", verifyOpts);
            Console.WriteLine("Cell A1 value after password change: " + verifyWb.Worksheets[0].Cells["A1"].Value);
        }
    }
}
