// Title: Convert an Excel workbook to separate CSV files while masking emails, phone numbers, and credit‑card numbers using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export each worksheet to its own CSV and redact email addresses with regex | How to mask phone numbers and credit‑card numbers in Excel cells before saving as CSV using Aspose.Cells | Apply custom regular‑expression patterns to cell values during workbook‑to‑CSV conversion in .NET | Iterate worksheets with Aspose.Cells and save them as CSV files after sensitive data masking | Aspose.Cells replace personal information in Excel sheets with asterisks during CSV export
// Tags: regex data masking Aspose.Cells | worksheet to CSV conversion Aspose.Cells C# | mask sensitive Excel cell values Aspose.Cells | export individual worksheets as CSV Aspose.Cells | redact personal information during Excel to CSV export

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace WorkbookToCsvMasking
{
    // The example loads an .xlsx workbook with Aspose.Cells, walks through every worksheet and each string cell, applies three regex masks (email, US phone, credit‑card) to hide sensitive data, updates cells when changes occur, and then saves each worksheet as a separate CSV file in a target folder.
    class Program
    {
        // Define regex patterns and their replacement masks
        private static readonly Dictionary<string, string> MaskPatterns = new Dictionary<string, string>
        {
            // Email addresses -> replace user part with asterisks
            { @"([a-zA-Z0-9._%+-]+)@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", "***@$2" },

            // US phone numbers (e.g., 123-456-7890, (123) 456-7890) -> mask all digits
            { @"\b(\+?\d{1,2}[\s-]?)?(\(?\d{3}\)?[\s-]?)?\d{3}[\s-]?\d{4}\b", "**********" },

            // Credit card numbers (simple 13-19 digit detection) -> mask all but last 4 digits
            { @"\b(\d{12,15})(\d{4})\b", "************$2" }
        };

        static void Main(string[] args)
        {
            // Input workbook path and output folder
            string inputPath = @"C:\Data\InputWorkbook.xlsx";
            string outputFolder = @"C:\Data\CsvOutput";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        Cells cells = sheet.Cells;

                        // Get the used range of the sheet (fully qualified to avoid ambiguity)
                        Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

                        int startRow = usedRange.FirstRow;
                        int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                        int startCol = usedRange.FirstColumn;
                        int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                        // Loop through each cell in the used range
                        for (int row = startRow; row <= endRow; row++)
                        {
                            for (int col = startCol; col <= endCol; col++)
                            {
                                Cell cell = cells[row, col];

                                // Process only string values
                                if (cell.Type == CellValueType.IsString && cell.Value != null)
                                {
                                    string original = cell.StringValue;
                                    string masked = ApplyMasking(original);
                                    // Update cell only if masking changed the value
                                    if (!original.Equals(masked))
                                    {
                                        cell.PutValue(masked);
                                    }
                                }
                            }
                        }

                        // Save the current worksheet as a CSV file (only active sheet is exported)
                        workbook.Worksheets.ActiveSheetIndex = sheet.Index;
                        string csvFileName = Path.Combine(outputFolder, $"{sheet.Name}.csv");
                        workbook.Save(csvFileName, SaveFormat.Csv);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                    }
                }

                Console.WriteLine("Workbook conversion and masking completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load or process workbook: {ex.Message}");
            }
        }

        // Applies all defined regex masks to the input text
        private static string ApplyMasking(string input)
        {
            string result = input;
            foreach (var kvp in MaskPatterns)
            {
                result = Regex.Replace(result, kvp.Key, kvp.Value);
            }
            return result;
        }
    }
}
