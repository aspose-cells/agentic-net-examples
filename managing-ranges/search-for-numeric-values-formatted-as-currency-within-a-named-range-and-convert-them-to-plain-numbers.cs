// Title: Convert Currency‑Formatted Cells in a Named Range to Numeric Values with Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses the named range "CurrencyRange", removes common currency symbols from string cells, parses them to numbers, converts numeric cells with currency styles to the General format, and saves the updated file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | currency to number conversion | remove currency symbols | named range processing | Excel cell style change | general number format | financial data cleaning
// Common Searches: Aspose.Cells convert currency formatted cells to numbers | remove $ € £ symbols from Excel cells using C# | change currency style to General in a named range Aspose.Cells | parse currency strings to numeric values in .NET Excel library | detect and replace built‑in currency formats with General format
// Developer Intent: Transform every cell inside the "CurrencyRange" that displays a currency—whether as a text string with a symbol or as a numeric value with a currency style—into a plain numeric value using the General format.
// Use Cases: Clean imported financial spreadsheets by stripping currency symbols before calculations. | Standardize data for CSV or database export when only raw numbers are required. | Prepare a reporting template where the designated range must contain unformatted numeric amounts.
// AI Prompts: Generate C# code with Aspose.Cells that locates a named range called 'CurrencyRange', removes any leading currency symbols from string cells, converts them to double, and sets the cell style to General. | Provide a method that scans an Aspose.Cells.Range, identifies cells using built‑in currency number formats (e.g., indices 164‑169) or custom formats containing a currency symbol, and changes those cells to the General number format.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCurrencyConversion
{
    // Loads an Excel workbook, accesses the named range "CurrencyRange", removes common currency symbols from string cells, parses them to numbers, converts numeric cells with currency styles to the General format, and saves the updated file using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "CurrencyRange"
                Name namedRange = workbook.Worksheets.Names["CurrencyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'CurrencyRange' not found.");
                    return;
                }

                // Get the actual cell range
                Aspose.Cells.Range range = namedRange.GetRange();

                // Process each cell in the range
                foreach (Cell cell in range)
                {
                    if (cell.Type == CellValueType.IsString)
                    {
                        string raw = cell.StringValue.Trim();

                        // Detect common currency symbols
                        if (raw.StartsWith("$") || raw.StartsWith("€") || raw.StartsWith("£") ||
                            raw.StartsWith("¥") || raw.StartsWith("₹"))
                        {
                            // Remove symbols and thousand separators
                            string cleaned = raw.Replace("$", "")
                                                .Replace("€", "")
                                                .Replace("£", "")
                                                .Replace("¥", "")
                                                .Replace("₹", "")
                                                .Replace(",", "")
                                                .Trim();

                            // Parse numeric part
                            if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                            {
                                cell.PutValue(numericValue);

                                // Set General number format
                                Style style = cell.GetStyle();
                                style.Number = 0; // General
                                cell.SetStyle(style);
                            }
                        }
                    }
                    else if (cell.Type == CellValueType.IsNumeric)
                    {
                        // Check if the cell uses a currency format
                        Style style = cell.GetStyle();
                        bool isCurrencyFormat = false;

                        // Common built‑in currency format indices (example values)
                        int[] currencyIndices = { 164, 165, 166, 167, 168, 169 };
                        foreach (int idx in currencyIndices)
                        {
                            if (style.Number == idx)
                            {
                                isCurrencyFormat = true;
                                break;
                            }
                        }

                        // Fallback: look for a currency symbol in a custom format string
                        if (!isCurrencyFormat && !string.IsNullOrEmpty(style.Custom) && style.Custom.Contains("$"))
                        {
                            isCurrencyFormat = true;
                        }

                        if (isCurrencyFormat)
                        {
                            // Change to General format
                            style.Number = 0;
                            cell.SetStyle(style);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
