// Title: Aspose.Cells C# – Convert Currency‑Formatted Numbers to General Format in a Named Range
// Description: Load an Excel workbook with Aspose.Cells, retrieve the named range "MyRange", iterate through its cells, detect numeric cells that use a currency style (custom format containing "$" or built‑in format ID 44), and change their style to the General number format (Number = 0) before saving the file.
// Keywords: Aspose.Cells | C# | .NET | currency format | remove currency formatting | named range | convert to general format | cell style | Excel number format | format ID 44 | custom number format | financial data cleaning
// Common Searches: Aspose.Cells change currency format to general | C# remove $ symbol from cells in named range | detect currency style Aspose.Cells | convert formatted numbers to plain numbers Aspose.Cells | iterate named range cells Aspose.Cells C#
// Developer Intent: Locate numeric cells formatted as currency inside a specific named range and convert them to the plain (General) number format.
// Use Cases: Sanitize financial reports by stripping currency symbols before exporting data to other systems. | Prepare data for calculations where currency formatting interferes with numeric operations. | Standardize worksheet appearance by converting currency‑styled cells in a defined range to General.
// AI Prompts: Provide C# code using Aspose.Cells to find numeric cells with a currency style in a named range and set their Number format to General while preserving other cell attributes. | Explain how to extend the sample to process multiple named ranges and support additional currency symbols such as € or £. | Suggest a technique to log the addresses of cells that were changed from currency to General during the conversion.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCurrencyConversion
{
    // Load an Excel workbook with Aspose.Cells, retrieve the named range "MyRange", iterate through its cells, detect numeric cells that use a currency style (custom format containing "$" or built‑in format ID 44), and change their style to the General number format (Number = 0) before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "MyRange"
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                    return;
                }

                // Get the actual range object
                Aspose.Cells.Range range = namedRange.GetRange();

                // Iterate through each cell in the range
                for (int row = range.FirstRow; row < range.FirstRow + range.RowCount; row++)
                {
                    for (int col = range.FirstColumn; col < range.FirstColumn + range.ColumnCount; col++)
                    {
                        Cell cell = workbook.Worksheets[0].Cells[row, col];

                        // Process only numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            // Retrieve the cell's current style
                            Style style = cell.GetStyle();

                            // Determine if the cell uses a currency format
                            bool isCurrencyFormat = false;

                            // Check custom format string for currency symbols
                            string customFormat = style.Custom;
                            if (!string.IsNullOrEmpty(customFormat) && customFormat.Contains("$"))
                            {
                                isCurrencyFormat = true;
                            }
                            else
                            {
                                // Check built‑in format IDs (44 is a common currency format)
                                if (style.Number == 44)
                                    isCurrencyFormat = true;
                            }

                            // If currency formatted, change to General format
                            if (isCurrencyFormat)
                            {
                                style.Number = 0; // General format
                                // Optionally clear custom format to avoid conflicts
                                style.Custom = string.Empty;
                                cell.SetStyle(style);
                            }
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
