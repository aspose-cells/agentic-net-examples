// Title: Mask emails, credit‑card numbers, and phone numbers in an Excel workbook and export to UTF‑8 CSV using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, scans all string cells, replaces email, credit‑card, and phone patterns using a dictionary of regular expressions, and saves the workbook as a UTF‑8 CSV with a custom separator. | Create a method that iterates over the used range of each worksheet, applies Regex.Replace for each sensitive pattern defined in a map, updates the cell values, and exports the result using TxtSaveOptions to a CSV file.
// Common Searches: how to replace personal data in Excel cells with placeholders before converting to CSV using Aspose.Cells C# | Aspose.Cells C# mask email addresses in workbook and save as CSV | C# export Excel to CSV while redacting credit card numbers with regular expressions | using Aspose.Cells to anonymize phone numbers in an .xlsx file and generate a CSV file | privacy filtering Excel data during CSV conversion with Aspose.Cells .NET
// Tags: Aspose.Cells mask sensitive data in Excel | C# regex redaction for workbook cells | TxtSaveOptions CSV export with UTF‑8 encoding | privacy filtering during Excel to CSV conversion | replace email credit‑card phone patterns using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an input.xlsx workbook with Aspose.Cells, iterates through all used string cells, applies regular‑expression masks for email addresses, credit‑card numbers, and phone numbers, updates the cell values, and then saves the workbook as a UTF‑8 CSV (output.csv) using TxtSaveOptions.
class WorkbookToCsvMasker
{
    // Define regex patterns and their replacement masks
    private static readonly Dictionary<string, string> SensitivePatterns = new Dictionary<string, string>
    {
        // Example: mask email addresses
        { @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", "[EMAIL]" },
        // Example: mask credit card numbers (simple pattern)
        { @"\b\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}\b", "[CREDIT_CARD]" },
        // Example: mask phone numbers
        { @"\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b", "[PHONE]" }
    };

    static void Main(string[] args)
    {
        // Input and output file paths
        string inputPath = "input.xlsx";   // Path to the source workbook
        string outputPath = "output.csv";  // Path for the resulting CSV

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and cells to mask sensitive data
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Use the used range for performance
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            string masked = MaskSensitiveData(original);
                            if (!original.Equals(masked))
                            {
                                cell.PutValue(masked);
                            }
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as CSV using TxtSaveOptions
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8,
                Separator = ','
            };
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved as CSV to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Applies all regex masks to the input text and returns the masked result
    private static string MaskSensitiveData(string input)
    {
        string result = input;
        foreach (var kvp in SensitivePatterns)
        {
            result = Regex.Replace(result, kvp.Key, kvp.Value, RegexOptions.Compiled);
        }
        return result;
    }
}
