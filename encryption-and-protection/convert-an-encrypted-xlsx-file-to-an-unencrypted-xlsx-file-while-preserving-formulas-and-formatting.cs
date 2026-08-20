// Title: Remove Password from an XLSX with Aspose.Cells (.NET) – Preserve Formulas & Formatting
// Description: Shows how to load a password‑protected workbook using Aspose.Cells LoadOptions, optionally unprotect the sheet structure, and save it as a plain XLSX file that keeps all formulas, styles, and other formatting intact.
// Keywords: Aspose.Cells decrypt XLSX | C# remove Excel password | load encrypted workbook Aspose | save unprotected workbook | preserve Excel formulas | Excel encryption .NET | Workbook.Unprotect | LoadOptions.Password
// Common Searches: asp.net strip password from excel file | c# decrypt encrypted xlsx using aspose.cells | open password protected excel and save without password | keep formulas when removing excel encryption | batch decrypt xlsx files Aspose.Cells
// Developer Intent: Open a secured Excel file and write a new version without encryption while leaving all content unchanged.
// Use Cases: Automated decryption of daily financial reports before data extraction | Pre‑processing step for Excel‑based ETL pipelines that require unprotected workbooks | Desktop utility that removes passwords from shared spreadsheets without altering formatting
// AI Prompts: Generate C# code with Aspose.Cells to open an encrypted XLSX, optionally unprotect the workbook structure, and save a plain copy preserving formulas and styles. | Explain the role of LoadOptions.Password in Aspose.Cells and demonstrate handling of an incorrect password exception. | Adapt the example to iterate over all .xlsx files in a folder, decrypt each one, and write the results to a target directory.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load a password‑protected workbook using Aspose.Cells LoadOptions, optionally unprotect the sheet structure, and save it as a plain XLSX file that keeps all formulas, styles, and other formatting intact.
    public class DecryptWorkbook
    {
        public static void Run()
        {
            // Path to the encrypted XLSX file
            string sourcePath = "encrypted.xlsx";

            // Path for the unencrypted output file
            string destPath = "decrypted.xlsx";

            // Password used to open the encrypted workbook
            string password = "yourPassword";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the encrypted workbook using LoadOptions with the password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // If the workbook structure is also protected, unprotect it (optional)
                // workbook.Unprotect(password);

                // Save the workbook without setting a password – this creates an unencrypted file
                workbook.Save(destPath, SaveFormat.Xlsx);

                Console.WriteLine($"Decryption completed. Unencrypted file saved to: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during decryption: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DecryptWorkbook.Run();
        }
    }
}
