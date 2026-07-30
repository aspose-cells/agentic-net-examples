// Title: Batch Decrypt Encrypted Excel (.xlsx) Files with a Shared Password using Aspose.Cells for .NET
// Description: Scans a folder for .xlsx files, loads each encrypted workbook with a common password via LoadOptions, optionally removes structure protection, clears the workbook password, and saves the unprotected copy to a target directory. Includes error handling for missing files and load failures.
// Keywords: Aspose.Cells batch decrypt | C# decrypt multiple Excel files | remove password from .xlsx programmatically | unprotect Excel workbook Aspose | load encrypted workbook with password | clear workbook Settings.Password | bulk Excel decryption .NET | automate Excel password removal | Aspose.Cells LoadOptions password | Excel file encryption removal C#
// Common Searches: how to batch decrypt encrypted Excel files using Aspose.Cells | C# remove password from multiple .xlsx files | Aspose.Cells bulk unprotect Excel workbooks | programmatically clear Excel file password .NET | load encrypted workbook with shared password Aspose
// Developer Intent: Decrypt a collection of password‑protected Excel workbooks and save them without any protection.
// Use Cases: Automated nightly job that strips passwords from incoming report spreadsheets before they are processed by analytics tools. | Bulk conversion of client‑submitted encrypted Excel files into unprotected versions for database import or data mining. | Server‑side decryption of archived Excel archives to make their contents searchable and editable.
// AI Prompts: Write C# code that uses Aspose.Cells to open every .xlsx file in a folder with a shared password, remove workbook and structure protection, clear Settings.Password, and save the files to another folder. | Explain how to catch and ignore exceptions when a workbook is not password‑protected while performing batch decryption with Aspose.Cells. | Provide a step‑by‑step tutorial for bulk decrypting encrypted Excel files and clearing the password property using Aspose.Cells for .NET. | Suggest performance optimizations for processing thousands of encrypted Excel files in a batch operation with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Scans a folder for .xlsx files, loads each encrypted workbook with a common password via LoadOptions, optionally removes structure protection, clears the workbook password, and saves the unprotected copy to a target directory. Includes error handling for missing files and load failures.
class BatchDecrypt
{
    static void Main()
    {
        // Folder containing encrypted Excel files
        string inputFolder = @"C:\EncryptedFiles";

        // Folder where unprotected files will be saved
        string outputFolder = @"C:\DecryptedFiles";

        // Shared password used for all encrypted workbooks
        string sharedPassword = "YourSharedPassword";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string sourcePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Verify the source file still exists
            if (!File.Exists(sourcePath))
                continue;

            try
            {
                // Load the encrypted workbook using the shared password
                LoadOptions loadOptions = new LoadOptions { Password = sharedPassword };
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // If the workbook structure is protected, unprotect it
                try
                {
                    workbook.Unprotect(sharedPassword);
                }
                catch
                {
                    // Ignore if the workbook is not protected
                }

                // Remove file‑level encryption by clearing the password property
                workbook.Settings.Password = null;

                // Save the unprotected workbook to the output folder
                string fileName = Path.GetFileName(sourcePath);
                string destPath = Path.Combine(outputFolder, fileName);
                workbook.Save(destPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to process '{sourcePath}': {ex.Message}");
            }
        }
    }
}
