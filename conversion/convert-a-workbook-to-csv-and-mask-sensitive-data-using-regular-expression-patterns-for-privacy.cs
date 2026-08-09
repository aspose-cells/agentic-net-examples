// Title: Mask Emails & Credit Card Numbers in Excel and Export to CSV with Aspose.Cells (C#)
// Description: Load an XLSX workbook, run regex patterns to replace email addresses and credit‑card numbers with placeholders, save a temporary masked file, convert it to CSV using Aspose.Cells.Utility.ConversionUtility, and clean up the intermediate file.
// Keywords: Aspose.Cells | C# | Excel to CSV conversion | mask sensitive data | regex replace in cells | PII redaction | email masking | credit card masking | data privacy | ConversionUtility
// Common Searches: Aspose.Cells replace email in Excel cells C# | How to mask credit card numbers in Excel before CSV export | Convert masked workbook to CSV using Aspose.Cells | Regex replace all cell values Aspose.Cells .NET | C# code to sanitize Excel data with Aspose.Cells
// Developer Intent: Redact personal identifiers in an Excel workbook with regex and generate a privacy‑safe CSV file.
// Use Cases: Prepare customer datasets for external sharing while complying with GDPR or PCI‑DSS. | Automate report pipelines that require PII removal before downstream analytics. | Create a one‑time masked export for auditors without altering the original workbook.
// AI Prompts: Write C# code that uses Aspose.Cells to replace phone numbers with [PHONE] in every string cell before exporting to CSV. | Show how to add a regex for Social Security numbers to the masking list and generate a pipe‑delimited text file with ConversionUtility. | Explain how to modify the example to keep the original workbook unchanged while producing a masked CSV.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Load an XLSX workbook, run regex patterns to replace email addresses and credit‑card numbers with placeholders, save a temporary masked file, convert it to CSV using Aspose.Cells.Utility.ConversionUtility, and clean up the intermediate file.
public class WorkbookToCsvMasking
{
    // Entry point
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Example usage
    public static void Run()
    {
        // Paths – adjust as needed
        string sourceFile = "input.xlsx";          // Original workbook
        string maskedFile = "masked.xlsx";         // Temporary masked workbook
        string outputCsv = "output.csv";           // Final CSV file

        // Verify source file exists to avoid FileNotFoundException
        if (!File.Exists(sourceFile))
        {
            Console.Error.WriteLine($"Source file '{sourceFile}' not found.");
            return;
        }

        // Load the source workbook (lifecycle: create & load)
        using (Workbook workbook = new Workbook(sourceFile))
        {
            // Define regex patterns and replacement values
            var patterns = new (string pattern, string replacement)[]
            {
                (@"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", "[EMAIL]"),
                (@"\b\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}\b", "[CREDIT_CARD]")
            };

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            string masked = original;

                            // Apply each regex pattern
                            foreach (var (pattern, replacement) in patterns)
                            {
                                masked = Regex.Replace(masked, pattern, replacement);
                            }

                            // If the value changed, write it back
                            if (!masked.Equals(original))
                            {
                                cell.PutValue(masked);
                            }
                        }
                    }
                }
            }

            // Save the masked workbook (lifecycle: save)
            workbook.Save(maskedFile, SaveFormat.Xlsx);
        }

        // Convert the masked workbook to CSV using the provided ConversionUtility rule
        if (File.Exists(maskedFile))
        {
            ConversionUtility.Convert(maskedFile, outputCsv);
        }
        else
        {
            Console.Error.WriteLine($"Masked file '{maskedFile}' was not created.");
            return;
        }

        // Optional: clean up the temporary masked file
        try
        {
            if (File.Exists(maskedFile))
            {
                File.Delete(maskedFile);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to delete temporary file '{maskedFile}': {ex.Message}");
        }

        Console.WriteLine($"Conversion completed. CSV saved to '{outputCsv}'.");
    }
}
