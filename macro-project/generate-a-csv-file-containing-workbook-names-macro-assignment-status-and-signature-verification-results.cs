// Title: Create a CSV inventory of Excel workbooks showing macro presence and digital signature status with Aspose.Cells for .NET
// Description: A C# console program that scans a set of Excel files, loads each workbook with Aspose.Cells, reads the file name, checks the HasMacro and IsDigitallySigned properties, safely escapes CSV fields, and writes a report containing the workbook name, macro flag (Yes/No) and signature flag (Yes/No).
// Keywords: Aspose.Cells CSV report | C# workbook macro detection | Excel digital signature check .NET | export workbook metadata to CSV | list Excel files macro status | Aspose.Cells HasMacro property | Aspose.Cells IsDigitallySigned | Excel security audit C#
// Common Searches: how to generate CSV of Excel files macro status using Aspose.Cells | C# list workbooks with digital signatures Aspose | export HasMacro and IsDigitallySigned to CSV | Aspose.Cells create inventory of macro-enabled workbooks | C# code to check Excel file signatures Aspose
// Developer Intent: Generate a CSV file that enumerates each workbook’s name together with Yes/No indicators for macro presence and digital signature verification.
// Use Cases: Security audit to identify macro‑enabled or unsigned Excel files before deployment. | Compliance reporting that shows which workbooks are digitally signed for governance purposes. | IT inventory generation for quick visibility of macro and signature status across a document library.
// AI Prompts: Write a C# method that scans a directory for .xlsx and .xlsm files and creates a CSV with workbook name, macro flag, and signature flag using Aspose.Cells. | Extend the sample to log errors to a separate file and add an "Error Message" column to the CSV output. | Add columns for the workbook's creation date and last modified date retrieved via Aspose.Cells and include them in the CSV report.

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that scans a set of Excel files, loads each workbook with Aspose.Cells, reads the file name, checks the HasMacro and IsDigitallySigned properties, safely escapes CSV fields, and writes a report containing the workbook name, macro flag (Yes/No) and signature flag (Yes/No).
class Program
{
    static void Main()
    {
        // Paths of the workbooks to be inspected
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsm",
            "SignedWorkbook.xlsx"
        };

        // Output CSV file path
        string csvOutputPath = "WorkbookReport.csv";

        try
        {
            // Create CSV and write header
            using (var writer = new StreamWriter(csvOutputPath))
            {
                writer.WriteLine("Workbook Name,Has Macro,Is Digitally Signed");

                foreach (var filePath in workbookFiles)
                {
                    // Verify that the file exists before attempting to load
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load workbook (no password supplied; will throw if protected)
                        using (Workbook wb = new Workbook(filePath))
                        {
                            // Ensure the FileName property reflects the actual file name
                            wb.FileName = Path.GetFileName(filePath);

                            string name = wb.FileName;
                            string hasMacro = wb.HasMacro ? "Yes" : "No";
                            string isSigned = wb.IsDigitallySigned ? "Yes" : "No";

                            // Write a CSV line for the current workbook
                            writer.WriteLine($"{EscapeCsv(name)},{hasMacro},{isSigned}");
                        }
                    }
                    catch (CellsException ex)
                    {
                        // Handle password-protected or corrupted files gracefully
                        Console.WriteLine($"Unable to process '{filePath}': {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        // Catch any other unexpected errors
                        Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"CSV report generated at: {csvOutputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to generate report: {ex.Message}");
        }
    }

    // Helper method to escape CSV fields containing commas, quotes, or new lines
    static string EscapeCsv(string field)
    {
        if (field.Contains("\"") || field.Contains(",") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
