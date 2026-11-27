// Title: How to strip currency formatting from numeric cells in a named range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove currency format from cells in a named range | convert currency formatted numbers to general format in Excel workbook using Aspose.Cells | detect currency number format indices 164-179 with Aspose.Cells .NET | clear custom currency number format symbols in Excel using Aspose.Cells C# | iterate over named range cells and change number style to General with Aspose.Cells
// Tags: remove currency number format Aspose.Cells | named range cell traversal Aspose.Cells C# | set cell style to General number format Aspose.Cells | identify built‑in currency format indices Aspose.Cells | clear custom currency format strings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, retrieves a named range, iterates through each cell in that range, detects numeric cells formatted as currency (both built‑in indices 164‑179 and custom formats containing currency symbols), clears the currency formatting by applying the General number style, and saves the updated workbook.
class CurrencyToNumberConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string namedRange = "MyNamedRange";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the named range
            AsposeRange range = workbook.Worksheets.GetRangeByName(namedRange);
            if (range == null)
            {
                Console.WriteLine($"Named range \"{namedRange}\" not found.");
                return;
            }

            // Determine the worksheet that contains the range
            Worksheet sheet = range.Worksheet;

            // Iterate through all cells in the range
            int firstRow = range.FirstRow;
            int firstCol = range.FirstColumn;
            int rowCount = range.RowCount;
            int colCount = range.ColumnCount;

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    Cell cell = sheet.Cells[firstRow + r, firstCol + c];

                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        Style style = cell.GetStyle();

                        // Determine if the cell is formatted as currency
                        bool isCurrency = false;

                        // Built‑in currency formats typically have indices 164‑179
                        if (style.Number >= 164 && style.Number <= 179)
                            isCurrency = true;

                        // Check custom format strings for currency symbols
                        if (!isCurrency && !string.IsNullOrEmpty(style.Custom))
                        {
                            string fmt = style.Custom;
                            if (fmt.Contains("$") || fmt.Contains("€") || fmt.Contains("£") || fmt.Contains("¥"))
                                isCurrency = true;
                        }

                        if (isCurrency)
                        {
                            // Remove currency formatting – set to General format
                            style.Number = 0;      // General format
                            style.Custom = null;   // Clear custom format
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
