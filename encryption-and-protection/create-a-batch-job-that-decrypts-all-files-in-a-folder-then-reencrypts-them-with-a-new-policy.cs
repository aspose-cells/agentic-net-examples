// Title: Batch decrypt and re‑encrypt Excel .xlsx files with a new password using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that scans a folder for encrypted .xlsx workbooks, opens each with a specified old password via Aspose.Cells LoadOptions, assigns a new workbook password, and saves the files to an output directory. | Write code that iterates over all Excel files in a source folder, removes the existing encryption by providing the old password, applies a new encryption policy, and writes the re‑encrypted files to a target folder using Aspose.Cells.
// Common Searches: aspnet change password for multiple encrypted Excel files using Aspose.Cells | c# script to decrypt and re‑encrypt .xlsx workbooks in a folder | bulk update Excel workbook encryption password Aspose.Cells .NET | load encrypted workbook with old password and save with new password Aspose.Cells example | automate re‑encryption of Excel files with new policy C#
// Tags: batch re‑encrypt Excel .xlsx files Aspose.Cells | load encrypted workbook with old password .NET | set new workbook password Aspose.Cells | process multiple encrypted Excel files C# | bulk update Excel encryption policy Aspose

using System;
using System.IO;
using Aspose.Cells;

// The C# console app enumerates .xlsx files in a source directory, loads each workbook with the old password using Aspose.Cells LoadOptions, sets a new password via workbook.Settings.Password, and saves the re‑encrypted files to a target folder.
class Program
{
    static void Main()
    {
        // Folder containing the encrypted Excel files
        string sourceFolder = @"C:\InputFolder";

        // Folder where re‑encrypted files will be saved
        string targetFolder = @"C:\OutputFolder";

        // Old encryption password (policy)
        string oldPassword = "oldPassword";

        // New encryption password (policy)
        string newPassword = "newPassword";

        try
        {
            // Ensure the source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the target folder exists
            Directory.CreateDirectory(targetFolder);

            // Process each .xlsx file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook using the old password
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = oldPassword
                    };
                    var workbook = new Workbook(filePath, loadOptions);

                    // Set the new password for the workbook
                    workbook.Settings.Password = newPassword;

                    // Determine the output file path
                    string fileName = Path.GetFileName(filePath);
                    string outputPath = Path.Combine(targetFolder, fileName);

                    // Save the workbook with the new encryption password
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Re‑encrypted file saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
