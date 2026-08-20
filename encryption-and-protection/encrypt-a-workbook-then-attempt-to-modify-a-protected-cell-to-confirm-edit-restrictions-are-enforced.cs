// Title: Encrypt an Excel workbook, protect a worksheet, and verify cell edit restrictions with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, writes a value to A1, protects the worksheet with a password, encrypts the file with a separate password, saves it, reloads using the file password, attempts to modify the locked cell (triggering an exception), then unprotects the sheet and updates the cell. Demonstrates workbook encryption, worksheet protection, and exception‑based validation of edit restrictions.
// Keywords: Aspose.Cells | C# | encrypt workbook | workbook password | worksheet protection | protect sheet Aspose.Cells | load encrypted Excel | LoadOptions password | verify cell protection | unprotect worksheet programmatically | exception handling Aspose.Cells | StrongCryptographicProvider | 128‑bit encryption
// Common Searches: Aspose.Cells encrypt workbook with password C# | How to protect a worksheet and prevent cell edits using Aspose.Cells | Load an encrypted Excel file and test sheet protection in .NET | Catch exception when modifying a protected cell with Aspose.Cells | Unprotect a worksheet after opening an encrypted workbook Aspose.Cells
// Developer Intent: Show how to apply file‑level encryption and sheet‑level protection with Aspose.Cells, then confirm that protected cells cannot be edited without the correct sheet password.
// Use Cases: Secure confidential reports by encrypting the file and locking all cells, then programmatically verify that unauthorized edits are blocked. | Integrate workbook protection into a document‑management system, ensuring that only users with the sheet password can modify critical data. | Automate regression tests that validate worksheet protection settings by attempting prohibited edits and checking for expected exceptions.
// AI Prompts: Generate C# code using Aspose.Cells to encrypt a workbook with a 256‑bit password, protect a worksheet, and assert that changing a locked cell throws an exception. | Explain step‑by‑step how Aspose.Cells enforces worksheet protection after opening an encrypted workbook with the correct file password. | Refactor the sample to log detailed exception information when a protected cell modification is blocked, and include unit‑test assertions.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionAndProtectionDemo
{
    // C# example that creates a workbook, writes a value to A1, protects the worksheet with a password, encrypts the file with a separate password, saves it, reloads using the file password, attempts to modify the locked cell (triggering an exception), then unprotects the sheet and updates the cell. Demonstrates workbook encryption, worksheet protection, and exception‑based validation of edit restrictions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Original Value");

            // Protect the worksheet with a password
            sheet.Protect(ProtectionType.All, "sheetPassword", null);

            // Set a password to encrypt the workbook file
            workbook.Settings.Password = "filePassword";
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted and protected workbook
            string filePath = "EncryptedProtectedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Load the workbook using the encryption password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "filePassword";
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Attempt to modify a protected cell without providing the worksheet password
            try
            {
                loadedSheet.Cells["A1"].PutValue("Attempted Modification");
                Console.WriteLine("Cell modified (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Modification blocked as expected: " + ex.Message);
            }

            // Unprotect the worksheet with the correct password and modify the cell
            loadedSheet.Unprotect("sheetPassword");
            loadedSheet.Cells["A1"].PutValue("Modified After Unprotect");
            Console.WriteLine("Cell value after unprotect: " + loadedSheet.Cells["A1"].StringValue);
        }
    }
}
