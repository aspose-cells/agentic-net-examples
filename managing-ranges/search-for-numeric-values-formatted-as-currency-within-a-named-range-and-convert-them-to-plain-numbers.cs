using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCurrencyConversion
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Retrieve the named range
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range \"MyRange\" not found.");
                    return;
                }

                // Obtain the actual range object (fully qualified to avoid ambiguity)
                Aspose.Cells.Range range = namedRange.GetRange();

                // Iterate through each cell in the range
                foreach (Cell cell in range)
                {
                    // Skip cells without a value
                    if (cell == null || cell.Value == null)
                        continue;

                    // Get the cell's number format (custom format string, if any)
                    Style cellStyle = cell.GetStyle();
                    string numberFormat = cellStyle.Custom;

                    // Get the cell's displayed string value
                    string stringValue = cell.StringValue?.Trim();

                    bool isCurrencyFormat = !string.IsNullOrEmpty(numberFormat) && numberFormat.Contains("$");
                    bool isCurrencyString = !string.IsNullOrEmpty(stringValue) && stringValue.StartsWith("$");

                    if (isCurrencyFormat || isCurrencyString)
                    {
                        // Parse the currency string to a numeric value
                        if (double.TryParse(stringValue, NumberStyles.Currency, CultureInfo.CurrentCulture, out double numericValue))
                        {
                            // Replace the cell value with the numeric value
                            cell.PutValue(numericValue);

                            // Reset the number format to General
                            cellStyle.Number = 0;          // Built‑in General format
                            cellStyle.Custom = null;       // Remove any custom format
                            cell.SetStyle(cellStyle);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}