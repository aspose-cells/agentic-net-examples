// Title: Create a timestamped log entry each time an Excel workbook's encryption is toggled with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to record encryption status changes in a log file | append timestamped entry when toggling Excel workbook password using Aspose.Cells | log workbook protection toggle events in .NET application | track encryption state changes of Excel files with Aspose.Cells | automate logging of Aspose.Cells workbook password updates
// Tags: Aspose.Cells encryption status audit | C# log Excel workbook password changes | timestamped protection log Aspose.Cells | automated workbook encryption tracking .NET | append encryption toggle entry to text file

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, checks its current encryption via Workbook.Settings.Password, toggles the encryption by setting or clearing the password, writes a UTC timestamped log entry describing the new encryption state to a text file, and saves the workbook.
class Program
{
    static void Main()
    {
        // Paths for input, output and log files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";
        string logPath = "encryption_log.txt";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(inputPath);

        // Check current encryption status
        bool isEncrypted = !string.IsNullOrEmpty(workbook.Settings.Password);

        // Toggle encryption status
        if (isEncrypted)
        {
            // Remove encryption
            workbook.Settings.Password = null;
        }
        else
        {
            // Apply encryption with a password
            workbook.Settings.Password = "SecretPassword123";
        }

        // Create a log entry for the toggle action
        string logEntry = $"{DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} - Encryption toggled. New status: {(workbook.Settings.Password != null ? "Encrypted" : "Not Encrypted")}";
        File.AppendAllText(logPath, logEntry + Environment.NewLine);

        // Save the workbook (save rule)
        workbook.Save(outputPath);
    }
}
