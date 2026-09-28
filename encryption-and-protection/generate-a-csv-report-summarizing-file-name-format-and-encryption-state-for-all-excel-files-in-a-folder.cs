// Title: Generate a CSV report of Excel file names, detected formats, and encryption status using Aspose.Cells in C#
// AI Prompts: Write a C# console application that scans a given directory, loads each Excel workbook with Aspose.Cells, captures the Workbook.FileFormat value, infers encryption by catching CellsException, and writes FileName, Format, and IsEncrypted columns to a CSV file. | Extend the program to accept a recursive‑scan flag, traverse subfolders, and add a WorkbookVersion column derived from the FileFormat enum. | Implement logging that records files failing to load for reasons other than encryption to a separate error log while continuing the CSV generation.
// Common Searches: C# Aspose.Cells create CSV list of Excel files with format and encryption flag | how to detect password protected Excel workbooks using Aspose.Cells .NET | list workbook file format for all .xls .xlsx files in a folder with Aspose.Cells | generate report of encrypted Excel files in a directory using C#
// Tags: Aspose.Cells enumerate workbook formats | detect encrypted Excel files Aspose.Cells | export workbook metadata to CSV C# | folder scan Excel files Aspose.Cells | CSV report of Excel file properties .NET

using System;
using System.IO;
using Aspose.Cells;

// Scans a specified folder for Excel files, loads each with Aspose.Cells to obtain the FileFormat enum, infers encryption by catching CellsException, and writes the filename, detected format, and encryption flag to a CSV report.
class Program
{
    static void Main(string[] args)
    {
        // Determine folder to scan – use first argument or current directory if none provided
        string folderPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder does not exist: {folderPath}");
            return;
        }

        // Path for the generated CSV report
        string csvPath = Path.Combine(folderPath, "ExcelFilesReport.csv");

        // Define Excel file extensions to consider
        string[] excelExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlsxml" };

        try
        {
            using (var writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("FileName,Format,IsEncrypted");

                // Enumerate files in the folder
                foreach (string filePath in Directory.GetFiles(folderPath))
                {
                    // Process only files with Excel extensions
                    if (Array.Exists(excelExtensions, ext => ext.Equals(Path.GetExtension(filePath), StringComparison.OrdinalIgnoreCase)))
                    {
                        bool isEncrypted = false;
                        string format = "Unknown";

                        try
                        {
                            // Load the workbook (Aspose.Cells automatically detects format)
                            Workbook workbook = new Workbook(filePath);

                            // Retrieve format (FileFormat property returns FileFormatType enum)
                            format = workbook.FileFormat.ToString();

                            // Aspose.Cells does not expose an IsEncrypted property directly.
                            // If loading succeeded, we assume the file is not encrypted.
                            isEncrypted = false;
                        }
                        catch (CellsException)
                        {
                            // If loading fails, assume the file is encrypted or corrupted
                            isEncrypted = true;
                        }
                        catch (Exception ex)
                        {
                            // Log unexpected errors but continue processing other files
                            Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                            continue;
                        }

                        // Write a line to the CSV
                        writer.WriteLine($"{Path.GetFileName(filePath)},{format},{isEncrypted}");
                    }
                }
            }

            Console.WriteLine($"CSV report generated at: {csvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to generate report: {ex.Message}");
        }
    }
}
