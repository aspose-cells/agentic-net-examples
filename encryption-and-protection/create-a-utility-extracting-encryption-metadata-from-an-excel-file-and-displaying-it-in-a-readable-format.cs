// Title: C# command‑line utility to extract Excel encryption and RMS protection metadata with Aspose.Cells
// Description: A lightweight C# console tool that accepts an Excel file path, uses Aspose.Cells FileFormatUtil to detect the file format, encryption status and RMS protection flag, prints the results, and attempts to load the workbook to show WorkbookSettings.IsEncrypted, handling password‑required exceptions gracefully.
// Keywords: Aspose.Cells encryption detection | C# Excel protection metadata | FileFormatUtil IsEncrypted | WorkbookSettings IsEncrypted | RMS protection Excel | command line Excel security tool | detect password protected workbook | Excel file format detection
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells .NET | retrieve RMS protection status from an Excel workbook | C# utility to display Excel encryption metadata | FileFormatUtil detect encrypted Excel file example | Aspose.Cells command line tool for workbook security
// Developer Intent: Create a simple command‑line application that reads an Excel file and reports its encryption and RMS protection details without requiring a password.
// Use Cases: Pre‑process incoming spreadsheets by confirming they are not password‑protected before batch operations. | Log encryption and RMS protection information for compliance audits of uploaded Excel files. | Provide a quick, human‑readable summary of workbook security settings in a file‑inspection utility.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect encryption and RMS protection of an Excel file and prints a friendly summary. | Write a method that returns a JSON object with FileFormatType, IsEncrypted, and IsProtectedByRMS for a given Excel path using Aspose.Cells. | Create robust error handling for loading an encrypted workbook without a password and capture the exception message for logging.

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptionMetadataUtility
{
    // A lightweight C# console tool that accepts an Excel file path, uses Aspose.Cells FileFormatUtil to detect the file format, encryption status and RMS protection flag, prints the results, and attempts to load the workbook to show WorkbookSettings.IsEncrypted, handling password‑required exceptions gracefully.
    public class Program
    {
        // Entry point of the utility
        public static void Main(string[] args)
        {
            // Expect a file path as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: EncryptionMetadataUtility <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Verify that the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found - {filePath}");
                return;
            }

            // Display encryption related metadata
            DisplayEncryptionMetadata(filePath);
        }

        // Extracts and prints encryption metadata for the specified Excel file
        public static void DisplayEncryptionMetadata(string filePath)
        {
            // Detect file format and retrieve metadata information
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            Console.WriteLine($"File: {Path.GetFileName(filePath)}");
            Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
            Console.WriteLine($"Is Protected By RMS: {formatInfo.IsProtectedByRMS}");

            // Provide a readable description based on the detection results
            if (formatInfo.IsEncrypted)
            {
                Console.WriteLine("The workbook is encrypted and requires a password to open.");
            }
            else
            {
                Console.WriteLine("The workbook is not encrypted.");
            }

            // Additional check using WorkbookSettings (optional, does not require saving)
            try
            {
                // Load the workbook without a password; if encrypted this will throw
                Workbook wb = new Workbook(filePath);
                Console.WriteLine($"WorkbookSettings.IsEncrypted: {wb.Settings.IsEncrypted}");
            }
            catch (Exception ex)
            {
                // Expected for encrypted files when password is not supplied
                Console.WriteLine($"Unable to load workbook without password: {ex.Message}");
            }
        }
    }
}
