// Title: Change the Password of an Encrypted Excel Workbook with Aspose.Cells for .NET (No Decryption)
// Description: Demonstrates how to load an already encrypted Excel file using LoadOptions, assign a new password via Settings.Password, and save the workbook so Aspose.Cells re‑encrypts it with the new password while leaving the worksheet data untouched.
// Keywords: Aspose.Cells change workbook password | update Excel password .NET | re‑encrypt Excel file programmatically | modify workbook encryption Aspose | change Excel file password without decrypting
// Common Searches: change password of encrypted Excel file Aspose.Cells | Aspose.Cells .NET update workbook password | re‑encrypt Excel workbook with new password | programmatically modify Excel file protection | Aspose.Cells load encrypted workbook password
// Developer Intent: Replace the existing password of an encrypted Excel workbook with a new one without exposing or decrypting its contents.
// Use Cases: Rotate passwords for archived reports to comply with security policies while keeping data encrypted. | Batch‑update workbook passwords across a document library without opening each file. | Migrate legacy protected workbooks to a new corporate password standard without data leakage.
// AI Prompts: Generate C# code using Aspose.Cells that changes the password of an already encrypted Excel workbook without reading its cell data. | Explain how Aspose.Cells re‑encrypts a workbook when Settings.Password is modified after loading with LoadOptions. | Show error‑handling patterns for loading an encrypted workbook with an incorrect password in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to load an already encrypted Excel file using LoadOptions, assign a new password via Settings.Password, and save the workbook so Aspose.Cells re‑encrypts it with the new password while leaving the worksheet data untouched.
class ChangeWorkbookPassword
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

        // Encrypt the workbook with the initial password
        wb.Settings.Password = "oldPassword";
        string encryptedPath = "EncryptedWorkbook.xlsx";
        wb.Save(encryptedPath);

        // Load the encrypted workbook using the original password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "oldPassword";
        Workbook loadedWb = new Workbook(encryptedPath, loadOptions);

        // Change the password without decrypting the content
        loadedWb.Settings.Password = "newPassword";

        // Save the workbook; it will be re‑encrypted with the new password
        loadedWb.Save("WorkbookWithNewPassword.xlsx");
    }
}
