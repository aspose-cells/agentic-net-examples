using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class WorkbookToCsvWithMasking
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Paths for the original workbook, the masked intermediate file, and the final CSV output
                string sourcePath = "input.xlsx";
                string maskedPath = "masked.xlsx";
                string csvPath = "output.csv";

                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the original workbook (lifecycle: create -> load)
                Workbook workbook = new Workbook(sourcePath);

                // Define regular expression patterns for sensitive data and their replacement masks
                var patterns = new Dictionary<string, string>
                {
                    // Email addresses
                    { @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", "[EMAIL_REDACTED]" },
                    // US phone numbers (e.g., 123-456-7890 or (123) 456-7890)
                    { @"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", "[PHONE_REDACTED]" },
                    // Credit card numbers (simple 16‑digit pattern)
                    { @"\b\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b", "[CC_REDACTED]" }
                };

                // Iterate through all used cells in the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Determine the used range to avoid scanning empty cells
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
                            foreach (var kvp in patterns)
                            {
                                masked = Regex.Replace(masked, kvp.Key, kvp.Value);
                            }

                            // If the value changed, write it back to the cell
                            if (!masked.Equals(original))
                            {
                                cell.PutValue(masked);
                            }
                        }
                    }
                }

                // Save the masked workbook (lifecycle: save)
                workbook.Save(maskedPath, SaveFormat.Xlsx);

                // Verify that the masked file was created before conversion
                if (!File.Exists(maskedPath))
                {
                    Console.WriteLine($"Masked file not created: {maskedPath}");
                    return;
                }

                // Convert the masked workbook to CSV using ConversionUtility
                ConversionUtility.Convert(maskedPath, csvPath);

                Console.WriteLine($"Workbook has been masked and converted to CSV: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}