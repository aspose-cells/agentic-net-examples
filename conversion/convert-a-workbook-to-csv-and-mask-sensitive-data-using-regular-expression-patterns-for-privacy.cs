// Title: Convert Excel to CSV with email & credit‑card masking using Aspose.Cells (C#)
// Description: Load an XLSX workbook with Aspose.Cells, scan string cells, replace email addresses and credit‑card numbers via Regex, save the sanitized workbook, and export it directly to CSV. Ideal for GDPR‑compliant data pipelines in .NET.
// Keywords: Aspose.Cells CSV export | C# Excel masking | regex data sanitization | email redaction Excel | credit card masking | privacy compliant CSV | Excel to CSV conversion .NET | GDPR data masking Aspose
// Common Searches: Aspose.Cells mask email before CSV export | C# replace credit card numbers in Excel workbook | how to sanitize Excel data with regex and convert to CSV | privacy compliant Excel to CSV conversion .NET | regex masking of PII in Aspose.Cells
// Developer Intent: Remove personally identifiable information from an Excel workbook and generate a clean CSV file using Aspose.Cells.
// Use Cases: Produce GDPR‑ready CSV reports by redacting customer emails. | Share financial worksheets without exposing credit‑card numbers. | Automate a data‑export workflow that strips PII before downstream processing.
// AI Prompts: Write C# code that uses Aspose.Cells to mask phone numbers with a [PHONE] placeholder before CSV conversion. | Extend the example to also replace URLs and social‑security numbers using regular expressions. | Explain how to stream a masked workbook directly to CSV with Aspose.Cells, eliminating the intermediate XLSX file.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // Load an XLSX workbook with Aspose.Cells, scan string cells, replace email addresses and credit‑card numbers via Regex, save the sanitized workbook, and export it directly to CSV. Ideal for GDPR‑compliant data pipelines in .NET.
    public class WorkbookToCsvWithMasking
    {
        public static void Run()
        {
            // Paths for the original workbook, the masked intermediate workbook, and the final CSV file
            string sourcePath = "input.xlsx";
            string maskedPath = "masked.xlsx";
            string csvPath = "output.csv";

            try
            {
                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the original workbook
                Workbook workbook = new Workbook(sourcePath);

                // Define regular expression patterns for sensitive data
                Regex emailRegex = new Regex(@"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", RegexOptions.IgnoreCase);
                Regex creditCardRegex = new Regex(@"\b\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}\b");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine the used range
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    // Scan each cell in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Process only string cells
                            if (cell.Type == CellValueType.IsString)
                            {
                                string original = cell.StringValue;
                                string masked = original;

                                // Mask email addresses
                                if (emailRegex.IsMatch(masked))
                                {
                                    masked = emailRegex.Replace(masked, "[EMAIL]");
                                }

                                // Mask credit card numbers
                                if (creditCardRegex.IsMatch(masked))
                                {
                                    masked = creditCardRegex.Replace(masked, "[CREDIT_CARD]");
                                }

                                // Update the cell only if a change occurred
                                if (!masked.Equals(original))
                                {
                                    cell.PutValue(masked);
                                }
                            }
                        }
                    }
                }

                // Save the workbook after masking (intermediate file)
                workbook.Save(maskedPath, SaveFormat.Xlsx);

                // Convert the masked workbook to CSV
                ConversionUtility.Convert(maskedPath, csvPath);

                Console.WriteLine($"Workbook has been masked and converted to CSV at: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToCsvWithMasking.Run();
        }
    }
}
