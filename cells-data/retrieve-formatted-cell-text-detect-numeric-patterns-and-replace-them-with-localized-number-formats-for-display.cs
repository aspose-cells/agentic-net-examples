using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsLocalizedNumberDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1234.56);          // Numeric value
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(7890);             // Integer value
            cells["A4"].PutValue("Note");
            cells["B4"].PutValue("Total: 9123.45"); // Text containing a number

            // Set the workbook culture to German (de-DE) for localized formatting
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Regular expression to detect numeric patterns in a string
            Regex numberRegex = new Regex(@"\d+([.,]\d+)?", RegexOptions.Compiled);

            // Iterate through the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Retrieve the formatted display string of the cell
                    string displayValue = cell.DisplayStringValue;

                    // Detect if the cell already contains a numeric value
                    if (cell.IsNumericValue)
                    {
                        // Apply a culture‑dependent custom number format
                        Style style = cell.GetStyle();
                        // "#,##0.00" will be rendered according to the workbook's CultureInfo (German uses comma as decimal separator)
                        style.CultureCustom = "#,##0.00";
                        cell.SetStyle(style);
                    }
                    else
                    {
                        // For text cells, check if they contain embedded numeric patterns
                        if (numberRegex.IsMatch(displayValue))
                        {
                            // Replace each numeric occurrence with a localized formatted version
                            string localized = numberRegex.Replace(displayValue, match =>
                            {
                                // Parse using invariant culture to avoid locale issues
                                if (double.TryParse(match.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                                {
                                    // Format using the workbook's culture (German in this example)
                                    return num.ToString("N", workbook.Settings.CultureInfo);
                                }
                                // If parsing fails, return the original match
                                return match.Value;
                            });

                            // Overwrite the cell with the new localized string (as plain text)
                            cell.PutValue(localized);
                        }
                    }
                }
            }

            // Output the final display strings to the console
            Console.WriteLine("Final cell display values after localization:");
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    Console.Write($"{cell.DisplayStringValue}\t");
                }
                Console.WriteLine();
            }

            // Save the workbook to verify the applied formats
            workbook.Save("LocalizedNumbersDemo.xlsx");
        }
    }
}