// Title: Convert Excel to CSV and Normalize Phone Numbers with Aspose.Cells (C#)
// Description: Loads an Excel workbook using Aspose.Cells, scans every used cell, detects phone‑number strings with a regular expression, rewrites them to the (XXX) XXX‑XXXX format, and saves the result as a CSV file.
// Keywords: Aspose.Cells CSV export C# | Excel to CSV conversion .NET | phone number formatting Aspose.Cells | regex phone normalization C# | standardize US phone numbers Excel | bulk data cleanup Aspose.Cells | C# workbook cell formatter
// Common Searches: How to format phone numbers in Excel before CSV export using Aspose.Cells | C# Aspose.Cells example to normalize phone numbers to (XXX) XXX-XXXX | Convert .xlsx to .csv and clean phone columns with Aspose.Cells | Regex based phone number standardization in Aspose.Cells workbook
// Developer Intent: Load an Excel file, reformat any phone‑number strings to a consistent pattern, and export the workbook as a CSV file.
// Use Cases: Prepare clean contact lists for CRM import. | Create uniform phone number columns for bulk SMS campaigns. | Generate CSV reports where phone fields must follow a specific US format.
// AI Prompts: Show a C# Aspose.Cells snippet that iterates all cells, detects phone numbers with regex, reformats them to (XXX) XXX-XXXX, and saves the workbook as CSV. | Explain how to extend the phone‑number regex to handle international formats while using Aspose.Cells. | Suggest performance tips for processing large worksheets when applying custom cell formatting before CSV export.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsPhoneNumberFormatter
{
    // Loads an Excel workbook using Aspose.Cells, scans every used cell, detects phone‑number strings with a regular expression, rewrites them to the (XXX) XXX‑XXXX format, and saves the result as a CSV file.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Load the workbook (create rule)
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define a regular expression that matches common phone number patterns
            // This pattern captures digits, optional separators, and optional country code
            Regex phoneRegex = new Regex(@"\+?(\d{1,3})?[\s\-\.]?\(?(\d{3})\)?[\s\-\.]?(\d{3})[\s\-\.]?(\d{4})");

            // Iterate over all used cells
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells that look like phone numbers
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue.Trim();

                        // Try to match the phone number pattern
                        Match match = phoneRegex.Match(original);
                        if (match.Success)
                        {
                            // Extract numeric groups (ignore country code if present)
                            string area = match.Groups[2].Value;
                            string prefix = match.Groups[3].Value;
                            string line = match.Groups[4].Value;

                            // Standardize to (XXX) XXX-XXXX format
                            string formatted = $"({area}) {prefix}-{line}";

                            // Replace the cell value with the formatted phone number
                            cell.PutValue(formatted);
                        }
                    }
                }
            }

            // Save the modified workbook as CSV (save rule)
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook converted to CSV with standardized phone numbers: {csvPath}");
        }
    }
}
