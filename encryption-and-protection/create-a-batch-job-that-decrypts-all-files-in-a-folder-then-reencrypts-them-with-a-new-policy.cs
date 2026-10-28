// Title: Batch decrypt and re‑encrypt Excel .xlsx files with a new password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells batch change password for multiple encrypted .xlsx files | How to load a password‑protected Excel workbook and save it with a different password using Aspose.Cells | Re‑encrypt all Excel files in a folder with a new password in .NET | Aspose.Cells process folder of encrypted workbooks and update password programmatically | Automate Excel file password update with LoadOptions in C#
// Tags: batch password update Aspose.Cells | decrypt and re‑encrypt .xlsx Aspose.Cells | load encrypted workbook LoadOptions | save workbook with new password | process folder of Excel files .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchEncryption
{
    // The BatchEncryptor class scans a specified directory for .xlsx files, opens each workbook with the existing password via LoadOptions, assigns a new password through Workbook.Settings.Password, and saves the workbook back, effectively decrypting and re‑encrypting every Excel file with the updated protection policy.
    class BatchEncryptor
    {
        /// <param name="folderPath">Path to the folder containing the Excel files.</param>
        /// <param name="oldPassword">Current password used to open the files.</param>
        /// <param name="newPassword">New password to apply to the files.</param>
        public static void ProcessFolder(string folderPath, string oldPassword, string newPassword)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files in the folder (adjust the pattern if needed)
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook using the old password (decryption)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = oldPassword
                    };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Apply the new password to the workbook
                    workbook.Settings.Password = newPassword;

                    // Overwrite the original file with the newly encrypted version
                    workbook.Save(filePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Successfully re‑encrypted: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }

        // Example usage
        static void Main(string[] args)
        {
            // Replace with actual values or retrieve from configuration
            string folder = @"C:\ExcelFiles";
            string oldPwd = "oldPassword123";
            string newPwd = "newSecurePassword456";

            try
            {
                ProcessFolder(folder, oldPwd, newPwd);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
