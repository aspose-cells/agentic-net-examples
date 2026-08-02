// Title: Detect Excel Format & Encryption, Decrypt if Needed, Process with Aspose.Cells (.NET)
// Description: A C# console example that checks whether an Excel file exists, uses Aspose.Cells FileFormatUtil to identify the workbook type and encryption flag, loads the file with LoadOptions.Password when protected, auto‑fits columns of the first worksheet, and saves the result as a new .xlsx file.
// Keywords: Aspose.Cells detect file format | Excel encryption detection .NET | load encrypted workbook Aspose.Cells | auto fit columns Aspose.Cells | save processed workbook C# | FileFormatUtil DetectFileFormat | LoadOptions.Password example | console app Excel processing
// Common Searches: how to check if an Excel file is password protected using Aspose.Cells | load encrypted Excel workbook with password in C# | detect Excel file type before opening with Aspose.Cells | auto‑fit columns after decrypting Excel file | Aspose.Cells example for encrypted workbook handling
// Developer Intent: Identify an Excel file’s format and encryption status, open it with the appropriate password if required, perform simple processing, and write the output to a new file.
// Use Cases: Batch‑process unknown Excel files by first detecting format and encryption, then applying the correct loading method. | Integrate format detection into a command‑line utility that accepts a file path and optional password. | Automate column auto‑fit for both protected and unprotected workbooks before converting them to .xlsx.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect an Excel file’s format and encryption flag, then opens it with a password if encrypted and auto‑fits the first worksheet. | Provide a robust error‑handling pattern for loading encrypted Excel files with Aspose.Cells, including missing password scenarios. | Create a PowerShell script that invokes a .NET console app (using Aspose.Cells) to process all Excel files in a directory, handling both encrypted and plain files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# console example that checks whether an Excel file exists, uses Aspose.Cells FileFormatUtil to identify the workbook type and encryption flag, loads the file with LoadOptions.Password when protected, auto‑fits columns of the first worksheet, and saves the result as a new .xlsx file.
    public class DetectAndDecryptDemo
    {
        /// <param name="filePath">Path to the Excel file.</param>
        /// <param name="password">Password to open the file if it is encrypted. Pass null or empty if unknown.</param>
        public static void Run(string filePath, string? password = null)
        {
            try
            {
                // Verify that the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File not found - {filePath}");
                    return;
                }

                // Detect file format and encryption status
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is encrypted: {formatInfo.IsEncrypted}");

                Workbook workbook;

                if (formatInfo.IsEncrypted)
                {
                    // Load encrypted workbook using the supplied password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                    {
                        // Empty string will cause an exception if password is required
                        Password = password ?? string.Empty
                    };
                    workbook = new Workbook(filePath, loadOptions);
                    Console.WriteLine("Workbook loaded with password.");
                }
                else
                {
                    // Load unencrypted workbook
                    workbook = new Workbook(filePath);
                    Console.WriteLine("Workbook loaded without password.");
                }

                // Example processing: auto‑fit columns of the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.AutoFitColumns();

                // Save the processed workbook to a new file
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(filePath) ?? string.Empty,
                    Path.GetFileNameWithoutExtension(filePath) + "_processed.xlsx");

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Processed workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for the console application
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("Usage: DetectAndDecryptDemo <excelFilePath> [password]");
                    return;
                }

                string filePath = args[0];
                string? password = args.Length > 1 ? args[1] : null;

                DetectAndDecryptDemo.Run(filePath, password);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
