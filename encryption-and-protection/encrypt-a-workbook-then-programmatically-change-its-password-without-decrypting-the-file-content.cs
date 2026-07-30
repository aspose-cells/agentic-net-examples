// Title: Change password of an encrypted Excel workbook with Aspose.Cells for .NET (C#)
// Description: Shows how to create an encrypted workbook, open it with LoadOptions using the current password, assign a new password via Settings.Password, and save the file so it remains encrypted—all without decrypting the worksheet data.
// Keywords: Aspose.Cells | C# | Excel encryption | change workbook password | password rotation | LoadOptions.Password | Settings.Password | encrypted .xlsx | programmatic password update | secure Excel file
// Common Searches: Aspose.Cells change Excel file password C# | Replace password of encrypted .xlsx using Aspose | Update password of encrypted workbook without decrypting | Load encrypted Excel with Aspose.Cells and set new password | Programmatic password rotation for Excel files .NET
// Developer Intent: Replace the existing protection password of an encrypted Excel workbook with a new one while keeping the file encrypted.
// Use Cases: Rotate passwords on archived reports to meet compliance policies. | Update workbook protection after a user changes their login credentials. | Automate batch password changes for multiple encrypted Excel files without exposing the data.
// AI Prompts: Write C# code that loads an encrypted .xlsx using Aspose.Cells, changes its password, and saves it with the new password. | Explain the role of LoadOptions.Password when opening an encrypted workbook and how Settings.Password influences the saved file. | Provide a step‑by‑step guide for rotating passwords on a collection of encrypted Excel files with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create an encrypted workbook, open it with LoadOptions using the current password, assign a new password via Settings.Password, and save the file so it remains encrypted—all without decrypting the worksheet data.
class ChangeWorkbookPassword
{
    static void Main()
    {
        // ---------- Create and encrypt the workbook ----------
        Workbook wb = new Workbook();                                   // create a new workbook
        wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");           // add some data
        wb.Settings.Password = "oldPass";                               // set initial password
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath);                                          // save encrypted file

        // ---------- Load the encrypted workbook ----------
        LoadOptions loadOpts = new LoadOptions();                       // create load options
        loadOpts.Password = "oldPass";                                  // provide current password
        Workbook loadedWb = new Workbook(encryptedPath, loadOpts);       // load the workbook

        // ---------- Change the password ----------
        loadedWb.Settings.Password = "newPass";                         // assign new password

        // ---------- Save the workbook with the new password ----------
        string newEncryptedPath = "encrypted_newpass.xlsx";
        loadedWb.Save(newEncryptedPath);                                // save with updated password
    }
}
