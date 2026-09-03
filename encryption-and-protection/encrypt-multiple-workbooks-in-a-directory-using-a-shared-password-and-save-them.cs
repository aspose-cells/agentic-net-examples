// Title: Encrypt multiple Excel workbooks in a folder with a shared password using Aspose.Cells for .NET
// AI Prompts: Create a C# program that scans a specified directory for .xlsx files, sets the same password on each workbook via Workbook.Settings.Password, and writes the encrypted files to a target folder. | Write a script using Aspose.Cells for .NET to load every Excel workbook in a source path, apply a common encryption password, and save the protected copies to a separate output directory.
// Common Searches: Aspose.Cells C# batch encrypt Excel files with same password | C# program to apply password protection to all .xlsx files in a folder | how to save encrypted workbooks to a different directory using Aspose.Cells | automate encryption of multiple Excel workbooks with Aspose.Cells .NET
// Tags: batch workbook encryption Aspose.Cells | Workbook.Settings.Password usage | encrypt .xlsx files C# Aspose.Cells | save protected Excel files to separate folder | shared password for multiple workbooks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example iterates over every .xlsx file in a source directory, loads each workbook with Aspose.Cells, assigns a common password through Workbook.Settings.Password, and saves the encrypted workbooks to a designated output folder.
class EncryptWorkbooks
{
    static void Main()
    {
        // Directory containing the source workbooks
        string sourceDirectory = @"C:\InputWorkbooks";

        // Directory where encrypted workbooks will be saved
        string destinationDirectory = @"C:\EncryptedWorkbooks";

        // Shared password for all workbooks
        string sharedPassword = "MySecurePassword";

        // Ensure the destination directory exists
        if (!Directory.Exists(destinationDirectory))
        {
            Directory.CreateDirectory(destinationDirectory);
        }

        // Process each .xlsx file in the source directory
        foreach (string sourceFilePath in Directory.GetFiles(sourceDirectory, "*.xlsx"))
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(sourceFilePath);

            // Apply password protection (encryption) to the workbook
            workbook.Settings.Password = sharedPassword;

            // Determine the output file path (overwrite with same name in destination folder)
            string fileName = Path.GetFileName(sourceFilePath);
            string destinationFilePath = Path.Combine(destinationDirectory, fileName);

            // Save the encrypted workbook
            workbook.Save(destinationFilePath);
        }

        Console.WriteLine("All workbooks have been encrypted and saved.");
    }
}
