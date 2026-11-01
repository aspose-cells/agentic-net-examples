// Title: Generate a CSV inventory of Excel files with format and password‑protection status using Aspose.Cells for .NET (C#)
// Common Searches: C# code to list Excel files in a folder and export their format and encryption flag to CSV with Aspose.Cells | How to detect password‑protected .xlsx files in a directory using the Aspose.Cells library | Create a batch report of workbook file types and protection status in .NET | Aspose.Cells load options to check if an Excel workbook is encrypted without providing a password | Generate CSV report of Excel file metadata (name, format, isEncrypted) in C#
// Tags: Aspose.Cells workbook format detection C# | enumerate Excel files encryption status Aspose.Cells | export Excel metadata to CSV .NET | detect password protected Excel workbook Aspose.Cells | batch processing Excel files Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Cells;

// The example scans a specified directory, filters supported Excel extensions, attempts to load each workbook with Aspose.Cells to determine its file format and whether it is password‑protected, and writes the filename, detected format, and encryption flag to a UTF‑8 CSV report.
class ExcelEncryptionReport
{
    static void Main()
    {
        // Folder containing the Excel files
        string folderPath = @"C:\ExcelFolder";

        // Verify the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Path for the generated CSV report
        string reportPath = Path.Combine(folderPath, "report.csv");

        // Prepare a StringBuilder for CSV content
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("FileName,Format,IsEncrypted");

        // Define Excel file extensions to look for
        string[] excelExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

        // Iterate through all files in the folder
        foreach (string filePath in Directory.GetFiles(folderPath))
        {
            string extension = Path.GetExtension(filePath);
            if (!excelExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
                continue; // Skip non‑Excel files

            // Ensure the file actually exists before attempting to load it
            if (!File.Exists(filePath))
                continue;

            string fileName = Path.GetFileName(filePath);
            string formatName = "Unknown";
            bool isEncrypted = false;

            try
            {
                // Attempt to load the workbook without a password
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Retrieve the detected format
                formatName = workbook.FileFormat.ToString();

                // If loading succeeded, the file is not encrypted
                isEncrypted = false;
            }
            catch (Exception ex)
            {
                // Check if the exception indicates a password‑protected file
                if (ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    isEncrypted = true;
                    formatName = extension.TrimStart('.').ToUpperInvariant();
                }
                else
                {
                    // Log unexpected errors and continue processing other files
                    Console.WriteLine($"Error processing '{fileName}': {ex.Message}");
                    continue;
                }
            }

            // Append the information as a CSV line
            csvBuilder.AppendLine($"{EscapeCsv(fileName)},{EscapeCsv(formatName)},{isEncrypted}");
        }

        // Write the CSV content to the report file
        try
        {
            File.WriteAllText(reportPath, csvBuilder.ToString(), Encoding.UTF8);
            Console.WriteLine($"Report generated at: {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write report: {ex.Message}");
        }
    }

    // Helper method to escape CSV fields containing commas or quotes
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
