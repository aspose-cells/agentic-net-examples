// Title: Mask SSN, credit‑card and email data with regex and export each worksheet to separate CSV files using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells mask sensitive data with regular expressions before CSV export | export each Excel worksheet to its own CSV file using Aspose.Cells .NET | replace SSN and credit‑card patterns in Excel cells with C# regex | how to configure TxtSaveOptions for UTF‑8 CSV output in Aspose.Cells | copy a worksheet to a temporary workbook and save as CSV with Aspose.Cells C#
// Tags: regex masking sensitive data Aspose.Cells | per‑worksheet CSV export Aspose.Cells | TxtSaveOptions CSV configuration C# | temporary workbook copy worksheet Aspose.Cells | mask SSN credit‑card email Excel C#

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, scans all string cells, and applies regex replacements to mask Social Security Numbers, 16‑digit credit‑card numbers, and email addresses. After masking, each worksheet is copied into a temporary workbook and saved as an individual UTF‑8 CSV file using TxtSaveOptions with a comma separator.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Define regex patterns and their replacement values for masking
            var maskPatterns = new (string Pattern, string Replacement)[]
            {
                // Social Security Number (e.g., 123-45-6789)
                (@"\b\d{3}-\d{2}-\d{4}\b", "XXX-XX-XXXX"),
                // Credit Card number (simple 16‑digit pattern)
                (@"\b\d{16}\b", "XXXXXXXXXXXXXXXX"),
                // Email address
                (@"\b[\w\.-]+@[\w\.-]+\.\w{2,4}\b", "masked@example.com")
            };

            // Iterate through every worksheet and every cell to apply masking
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only string cells; numeric/date cells are left unchanged
                    if (cell.Type == CellValueType.IsString)
                    {
                        string originalText = cell.StringValue;
                        string maskedText = originalText;

                        // Apply each regex replacement
                        foreach (var (Pattern, Replacement) in maskPatterns)
                        {
                            maskedText = Regex.Replace(maskedText, Pattern, Replacement);
                        }

                        // Write back the masked value if it changed
                        if (!maskedText.Equals(originalText))
                        {
                            cell.PutValue(maskedText);
                        }
                    }
                }
            }

            // Export each worksheet as an individual CSV file
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                try
                {
                    // Create a temporary workbook that contains only the current sheet
                    Workbook tempWorkbook = new Workbook();

                    // Copy the masked sheet into the temporary workbook
                    int copiedIndex = tempWorkbook.Worksheets.AddCopy(sheet.Name);
                    Worksheet copiedSheet = tempWorkbook.Worksheets[copiedIndex];

                    // Remove the default empty sheet that Aspose.Cells creates
                    if (tempWorkbook.Worksheets.Count > 1)
                    {
                        tempWorkbook.Worksheets.RemoveAt(0);
                    }

                    // Configure CSV save options using TxtSaveOptions (compatible with CSV)
                    TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                    {
                        Separator = ',',                     // Use comma as delimiter
                        Encoding = Encoding.UTF8             // UTF‑8 encoding
                    };

                    // Build the CSV file name: originalname_sheetname.csv
                    string csvFileName = $"{Path.GetFileNameWithoutExtension(inputPath)}_{sheet.Name}.csv";

                    // Save the temporary workbook as CSV
                    tempWorkbook.Save(csvFileName, csvOptions);
                    Console.WriteLine($"Saved CSV: {csvFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
