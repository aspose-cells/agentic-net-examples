using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsPhoneNumberFormatter
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel workbook
                string sourcePath = "input.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: Source file \"{sourcePath}\" not found.");
                    return;
                }

                // Load the workbook (creation rule)
                Workbook workbook = new Workbook(sourcePath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define a regular expression to remove non‑digit characters
                Regex digitsOnly = new Regex(@"\D");

                // Assume phone numbers are in column B (index 1), starting from row 0
                int firstRow = 0;
                int lastRow = cells.MaxDataRow; // Last row that contains data

                for (int row = firstRow; row <= lastRow; row++)
                {
                    Cell cell = cells[row, 1]; // Column B

                    // Get the raw string value (unformatted)
                    string raw = cell.StringValue?.Trim();

                    // Process only if the cell contains a non‑empty value
                    if (!string.IsNullOrEmpty(raw))
                    {
                        // Remove all non‑digit characters
                        string digits = digitsOnly.Replace(raw, "");

                        // If we have exactly 10 digits, format as (###) ###‑####
                        if (digits.Length == 10)
                        {
                            string formatted = $"({digits.Substring(0, 3)}) {digits.Substring(3, 3)}-{digits.Substring(6, 4)}";

                            // Replace the cell value with the formatted phone number
                            cell.PutValue(formatted);

                            // Apply a custom number format (display only)
                            Style style = cell.GetStyle();
                            style.Custom = "(###) ###-####";
                            cell.SetStyle(style);
                        }
                    }
                }

                // Save the workbook as CSV (save rule)
                string csvPath = "output.csv";

                // Ensure the directory for the CSV exists
                string csvDir = Path.GetDirectoryName(csvPath);
                if (!string.IsNullOrEmpty(csvDir) && !Directory.Exists(csvDir))
                {
                    Directory.CreateDirectory(csvDir);
                }

                workbook.Save(csvPath, SaveFormat.Csv);

                Console.WriteLine($"Workbook converted to CSV and saved at: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}