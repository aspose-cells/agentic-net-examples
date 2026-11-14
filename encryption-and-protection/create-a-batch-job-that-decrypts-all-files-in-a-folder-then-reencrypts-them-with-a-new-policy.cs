// Title: How to batch decrypt and re‑encrypt Excel .xlsx workbooks with a new password using Aspose.Cells for .NET (C#)
// Common Searches: C# batch process encrypted .xlsx files to change password with Aspose.Cells | Aspose.Cells load password‑protected workbook and save with new encryption policy in a folder | How to programmatically re‑encrypt multiple Excel files using a new password in .NET | Decrypt and re‑encrypt Excel files in bulk using Aspose.Cells LoadOptions and Workbook.Settings | Automate changing workbook passwords for all .xlsx files in a directory with Aspose.Cells
// Tags: batch decrypt reencrypt Excel Aspose.Cells | load protected xlsx with LoadOptions C# | set workbook password Aspose.Cells | save workbook with new encryption password | process Excel files in directory Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# utility scans a source folder for .xlsx files, loads each encrypted workbook using LoadOptions with the old password, assigns a new password via Workbook.Settings.Password, and saves the files to an output folder, handling errors per file and logging progress.
class BatchDecryptReencrypt
{
    // Folder containing the encrypted Excel files
    private const string InputFolder = @"C:\EncryptedFiles";

    // Folder where re‑encrypted files will be saved
    private const string OutputFolder = @"C:\ReEncryptedFiles";

    // Current password used to decrypt the files
    private const string OldPassword = "oldPassword123";

    // New password (policy) to apply when re‑encrypting
    private const string NewPassword = "NewStrongPassword!@#";

    static void Main()
    {
        try
        {
            // Verify input folder exists
            if (!Directory.Exists(InputFolder))
            {
                Console.WriteLine($"Input folder not found: {InputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(OutputFolder);

            // Get all Excel files in the input folder (top‑level only)
            string[] files = Directory.GetFiles(InputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                try
                {
                    // Load the workbook using the old password
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = OldPassword
                    };
                    var workbook = new Workbook(filePath, loadOptions);

                    // Prepare the output file path
                    string fileName = Path.GetFileName(filePath);
                    string outputPath = Path.Combine(OutputFolder, fileName);

                    // Set the new password for encryption when saving
                    workbook.Settings.Password = NewPassword;

                    // Save the workbook with the new encryption policy
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed: {fileName}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch job completed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
