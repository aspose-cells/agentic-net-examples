// Title: Decrypt a password‑protected Excel workbook and save it unencrypted using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to verify the existence of an encrypted XLSX file, configure LoadOptions with the workbook password, load the file with Aspose.Cells, read a cell to confirm access, clear the password setting, and write a new copy without encryption.
// Keywords: Aspose.Cells | C# load encrypted Excel | LoadOptions password | remove workbook password | decrypt XLSX .NET | save unprotected workbook | Excel encryption Aspose
// Common Searches: Open encrypted Excel file with Aspose.Cells C# | How to remove password from .xlsx using Aspose | Aspose.Cells LoadOptions example for protected workbook | Decrypt Excel workbook programmatically .NET | Save Excel without password Aspose.Cells
// Developer Intent: Open a protected workbook, read its contents, strip the password, and write a plain version.
// Use Cases: Read confidential financial data from a secured spreadsheet, then generate an unprotected version for downstream reporting. | Automate batch decryption of multiple password‑locked Excel files before archiving them in a plain‑text repository. | Validate cell values in a protected workbook before removing its encryption to ensure data integrity.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions to open an encrypted .xlsx with a given password and saves it without any protection. | Explain the steps to clear a workbook's password by setting Workbook.Settings.Password to null and then persisting the file. | Provide a pattern for iterating over a folder of password‑protected Excel files, removing each password, and logging success or failure for every file.

using System;
using System.IO;
using Aspose.Cells;

namespace LoadEncryptedWorkbookDemoApp
{
    // Demonstrates how to verify the existence of an encrypted XLSX file, configure LoadOptions with the workbook password, load the file with Aspose.Cells, read a cell to confirm access, clear the password setting, and write a new copy without encryption.
    public class LoadEncryptedWorkbookDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Path to the encrypted workbook file
            string encryptedFilePath = "encrypted.xlsx";

            // Verify the encrypted file exists
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"File not found: {encryptedFilePath}");
                return;
            }

            // Password used to protect the workbook
            string password = "mySecret";

            // Create LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.Password = password;

            // Load the password‑protected workbook using the load options
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Example verification: output the value of cell A1 from the first worksheet
            Console.WriteLine("Loaded workbook. Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].Value);

            // Remove the password protection from the workbook
            workbook.Settings.Password = null;

            // Save the workbook without password protection
            string unprotectedFilePath = "unprotected.xlsx";
            workbook.Save(unprotectedFilePath);

            Console.WriteLine("Workbook saved without password to: " + unprotectedFilePath);
        }
    }
}
