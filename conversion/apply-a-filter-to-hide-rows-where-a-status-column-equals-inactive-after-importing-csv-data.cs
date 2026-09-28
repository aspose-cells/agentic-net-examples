// Title: Hide rows with 'Inactive' status after loading a CSV into an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a CSV file with Aspose.Cells, detects the column named 'Status', and sets Row.IsHidden = true for every row where the cell value equals 'Inactive'. | Show how to apply a row‑level hide filter in an Aspose.Cells workbook after converting CSV data, including handling a missing 'Status' header and saving the result as XLSX.
// Common Searches: Aspose.Cells C# hide rows where Status column equals Inactive after CSV import | How to programmatically hide rows based on a column value in a workbook created from CSV using Aspose.Cells | C# Aspose.Cells hide rows with specific cell text before saving to XLSX | Filter out inactive records in Excel by hiding rows with Aspose.Cells .NET
// Tags: Row.IsHidden Aspose.Cells example | LoadOptions CSV Aspose.Cells .NET | status column detection Aspose.Cells | hide rows based on cell value C# | convert CSV to XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads a CSV file into a workbook, locates the 'Status' column, and hides rows where the cell value equals 'Inactive' before saving the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load CSV data into a new workbook using appropriate LoadOptions
            var loadOptions = new LoadOptions(LoadFormat.Csv);
            var workbook = new Workbook(inputPath, loadOptions);

            // Reference the first worksheet
            var sheet = workbook.Worksheets[0];

            // Find the column index that contains the "Status" header
            int statusCol = -1;
            for (int c = 0; c <= sheet.Cells.MaxColumn; c++)
            {
                if (sheet.Cells[0, c].StringValue.Equals("Status", StringComparison.OrdinalIgnoreCase))
                {
                    statusCol = c;
                    break;
                }
            }
            // Default to column B (index 1) if header not found
            if (statusCol == -1) statusCol = 1;

            // Hide rows where the status equals "Inactive"
            int firstDataRow = 1; // assuming row 0 is the header
            int lastDataRow = sheet.Cells.MaxDataRow;
            for (int r = firstDataRow; r <= lastDataRow; r++)
            {
                try
                {
                    string cellValue = sheet.Cells[r, statusCol].StringValue;
                    if (cellValue.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                    {
                        // Hide the entire row using RowCollection
                        var row = sheet.Cells.Rows[r];
                        row.IsHidden = true; // Correct property to hide a row
                    }
                }
                catch (Exception exRow)
                {
                    Console.WriteLine($"Error processing row {r}: {exRow.Message}");
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
