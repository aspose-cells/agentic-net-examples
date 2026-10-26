// Title: Batch toggle Excel workbook encryption and log each change with Aspose.Cells for .NET
// Common Searches: how to programmatically enable or disable password protection on multiple .xlsx files using Aspose.Cells in C# | c# code to toggle encryption of Excel workbooks and write a timestamped log entry | process a folder of Excel files, change encryption status and record actions with Aspose.Cells | Aspose.Cells load encrypted workbook with known password and remove protection | log encryption enable/disable events for Excel files in a .NET application
// Tags: batch change Excel password protection Aspose.Cells | Aspose.Cells load encrypted workbook with password | Aspose.Cells remove workbook password | Aspose.Cells save workbook with password | C# log workbook encryption actions

using System;
using System.IO;
using Aspose.Cells;

// The example scans a directory for .xlsx files, attempts to open each workbook using a known password, detects whether it is currently encrypted, toggles the password protection (adds or removes the password), saves the modified workbook to an output folder, and appends a UTC timestamped entry to a log file indicating whether encryption was enabled or disabled, while handling missing files and load/save errors.
class WorkbookEncryptionProcessor
{
    static void Main(string[] args)
    {
        try
        {
            // Folder containing workbooks to process
            string inputFolder = @"C:\Input";
            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\Output";
            // Path to the log file
            string logFile = Path.Combine(outputFolder, "encryption_log.txt");

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                ToggleEncryptionAndLog(filePath, outputFolder, logFile);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    static void ToggleEncryptionAndLog(string inputPath, string outputFolder, string logPath)
    {
        // Verify the input file exists
        if (!File.Exists(inputPath))
        {
            File.AppendAllText(logPath,
                $"{DateTime.UtcNow:u} | File not found: {inputPath}{Environment.NewLine}");
            return;
        }

        const string knownPassword = "Secret123";
        Workbook workbook = null;
        bool wasEncrypted = false;

        try
        {
            // Attempt to load assuming the file is encrypted with the known password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = knownPassword
            };
            workbook = new Workbook(inputPath, loadOptions);
            wasEncrypted = true;
        }
        catch (Exception)
        {
            try
            {
                // Load without a password (assumes the file is not encrypted)
                workbook = new Workbook(inputPath);
                wasEncrypted = false;
            }
            catch (Exception ex)
            {
                // Log load failure and exit
                string fileName = Path.GetFileName(inputPath);
                string errorEntry = $"{DateTime.UtcNow:u} | File: {fileName} | Load Error: {ex.Message}";
                File.AppendAllText(logPath, errorEntry + Environment.NewLine);
                return;
            }
        }

        try
        {
            // Toggle encryption status
            if (wasEncrypted)
            {
                // Remove encryption
                workbook.Settings.Password = null;
            }
            else
            {
                // Enable encryption with the predefined password
                workbook.Settings.Password = knownPassword;
            }

            // Save the workbook to the output folder with the same file name
            string fileName = Path.GetFileName(inputPath);
            string outputPath = Path.Combine(outputFolder, fileName);
            workbook.Save(outputPath);

            // Log the action
            string action = wasEncrypted ? "Disabled" : "Enabled";
            string logEntry = $"{DateTime.UtcNow:u} | File: {fileName} | Encryption {action}";
            File.AppendAllText(logPath, logEntry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // Log any errors that occur during processing
            string fileName = Path.GetFileName(inputPath);
            string errorEntry = $"{DateTime.UtcNow:u} | File: {fileName} | Error: {ex.Message}";
            File.AppendAllText(logPath, errorEntry + Environment.NewLine);
        }
    }
}
