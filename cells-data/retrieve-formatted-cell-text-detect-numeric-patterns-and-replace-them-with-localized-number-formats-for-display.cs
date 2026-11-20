// Title: Replace numeric substrings in Excel cell display text with culture-specific formatting using Aspose.Cells for .NET
// Common Searches: how to format numbers inside Excel cell strings according to the current culture with Aspose.Cells C# | C# Aspose.Cells example for detecting and localizing numeric patterns in displayed cell text | replace integer and decimal values in Excel cell text with localized format using Aspose.Cells | using regex to find numbers in Excel cell display text and apply culture-specific formatting in .NET | Aspose.Cells C# code to convert embedded numbers in cell strings to locale-aware format
// Tags: culture-specific number formatting in Excel cells Aspose.Cells | regex detection of numeric substrings in cell text C# | replace numbers in displayed cell strings Aspose.Cells | localize embedded numeric values Excel .NET | format numbers within cell string values using Aspose.Cells | process used range cell text for localization C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsNumberLocalization
{
    // The sample loads a workbook, iterates over the used range of the first worksheet, reads each cell's displayed string, uses a regular expression to locate integer or decimal numbers, parses them with invariant culture, formats them with the current culture's numeric format, replaces the original substrings, writes the updated string back to the cell, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths (use arguments or defaults)
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Regex to find numeric patterns (integers or decimals, optional sign)
                Regex numberRegex = new Regex(@"-?\d+(\.\d+)?", RegexOptions.Compiled);

                // Get the used range of the worksheet (fully qualified to avoid ambiguity)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Iterate through all used cells
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Retrieve the displayed text of the cell
                        string formattedText = cell.StringValue;

                        // Find numeric substrings within the text
                        MatchCollection matches = numberRegex.Matches(formattedText);
                        if (matches.Count == 0)
                            continue; // No numbers to localize

                        string localizedText = formattedText;

                        // Replace each numeric substring with a culture‑specific format
                        foreach (Match match in matches)
                        {
                            if (double.TryParse(match.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
                            {
                                string localizedNumber = number.ToString("N", CultureInfo.CurrentCulture);
                                localizedText = localizedText.Replace(match.Value, localizedNumber);
                            }
                        }

                        // Update the cell with the localized string
                        cell.PutValue(localizedText);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
