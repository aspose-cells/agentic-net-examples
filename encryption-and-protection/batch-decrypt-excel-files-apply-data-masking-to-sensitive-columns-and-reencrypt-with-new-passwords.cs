// Title: C# – Batch Decrypt, Mask Sensitive Columns, and Re‑Encrypt Excel Workbooks with Aspose.Cells
// Description: A .NET utility that recursively scans a directory, opens password‑protected Excel files using the old password, replaces values in chosen columns with a mask, applies a new workbook password, and saves each file, handling Aspose.Cells exceptions and logging errors.
// Keywords: Aspose.Cells C# | batch decrypt Excel | mask sensitive data in spreadsheets | re‑encrypt workbook password | bulk Excel processing .NET | Excel encryption detection | LoadOptions password Aspose | Workbook.Unprotect example | GDPR data redaction Excel | automated Excel security update
// Common Searches: how to decrypt multiple Excel files with Aspose.Cells | C# code to mask columns in encrypted workbooks | re‑protect Excel spreadsheets after data redaction | bulk processing of password‑protected Excel files .NET | replace values in specific Excel columns programmatically
// Developer Intent: Programmatically open encrypted Excel workbooks, replace data in designated columns with a placeholder, and save the files encrypted under a new password.
// Use Cases: Sanitizing personally identifiable information before sharing archived spreadsheets. | Automating GDPR‑compliant redaction of SSN or credit‑card columns across a repository of protected files. | Applying a corporate‑wide password change after masking sensitive data in bulk.
// AI Prompts: Write C# code using Aspose.Cells to open every encrypted Excel file in a folder, mask columns 0 and 2 with '*****', and save each file with a new password. | Explain the difference between workbook encryption and worksheet protection when processing Excel files with Aspose.Cells. | Refactor the loop to add structured logging (file name, status, errors) and skip non‑Excel formats efficiently.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A .NET utility that recursively scans a directory, opens password‑protected Excel files using the old password, replaces values in chosen columns with a mask, applies a new workbook password, and saves each file, handling Aspose.Cells exceptions and logging errors.
class BatchDecryptMaskEncrypt
{
    static void Main()
    {
        // Folder containing Excel files
        string folderPath = @"C:\ExcelFiles";

        // Old password used to decrypt existing protected files
        const string oldPassword = "oldPass123";

        // New password to encrypt the files after masking
        const string newPassword = "newPass456";

        // Zero‑based column indexes that contain sensitive data (e.g., 0 = A, 2 = C)
        int[] sensitiveColumns = new int[] { 0, 2 };

        // Mask string to replace sensitive values
        const string maskValue = "*****";

        // Verify the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Process each Excel file in the folder (including subfolders)
        foreach (string filePath in Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories))
        {
            // Consider only Excel formats
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                continue;

            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Detect if the file is encrypted
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                bool isEncrypted = formatInfo.IsEncrypted;

                // Prepare load options
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                if (isEncrypted)
                    loadOptions.Password = oldPassword; // use old password for encrypted files

                // Load the workbook
                Workbook workbook = new Workbook(filePath, loadOptions);

                // If the workbook was encrypted, unprotect it (if protection is set)
                if (isEncrypted)
                {
                    try
                    {
                        workbook.Unprotect(oldPassword);
                    }
                    catch (CellsException)
                    {
                        // Ignore if the workbook is not protected with a password
                    }
                }

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Apply masking only to the specified sensitive columns
                    foreach (int colIndex in sensitiveColumns)
                    {
                        if (colIndex > maxCol) continue; // skip if column is beyond used range

                        for (int row = 0; row <= maxRow; row++)
                        {
                            // Replace the cell value with the mask
                            sheet.Cells[row, colIndex].PutValue(maskValue);
                        }
                    }
                }

                // Set new password for the workbook (this encrypts the file on save)
                workbook.Settings.Password = newPassword;

                // Save the workbook back to the same file (overwrites original)
                workbook.Save(filePath);
            }
            catch (CellsException ex)
            {
                // Handle Aspose.Cells specific errors (e.g., invalid password) and continue with next file
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.WriteLine($"Unexpected error for file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
