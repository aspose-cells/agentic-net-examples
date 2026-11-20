// Title: Extract displayed currency strings from column A of an Excel workbook and write parsed numeric values to column B using Aspose.Cells for .NET
// Common Searches: how to retrieve the displayed text of a currency‑formatted cell with Aspose.Cells in C# | parse Excel currency strings to double and place results in another column using Aspose.Cells | Aspose.Cells read cell StringValue and convert to numeric value for calculations in C# | convert formatted currency cells to numbers for further processing with Aspose.Cells .NET | save workbook after updating cells with parsed currency values using Aspose.Cells
// Tags: read cell StringValue Aspose.Cells | parse currency string .NET | write numeric value to adjacent cell Aspose.Cells | process column A rows Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook with Aspose.Cells, iterates through all used rows in column A, reads each cell's displayed string via StringValue, attempts to parse it as a currency using the current culture, writes the resulting double (or zero on failure) into column B of the same row, and saves the updated workbook to output.xlsx while handling missing files and I/O errors.
class CurrencyParser
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook; if it is password‑protected, an exception will be thrown
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the range to process (column A, from first row to the last used row)
        int firstRow = 0;                     // zero‑based index (row 1)
        int lastRow = cells.MaxDataRow;       // last row that contains data

        for (int row = firstRow; row <= lastRow; row++)
        {
            // Cell containing the formatted currency string (column A)
            Cell currencyCell = cells[row, 0];

            // Retrieve the displayed string (includes currency symbol, grouping, etc.)
            string formattedValue = currencyCell.StringValue;

            // Try to parse the formatted currency string to a numeric value
            if (double.TryParse(formattedValue, NumberStyles.Currency, CultureInfo.CurrentCulture, out double numericValue))
            {
                // Store the numeric value in the adjacent cell (column B)
                cells[row, 1].PutValue(numericValue);
            }
            else
            {
                // Parsing failed – write zero (or leave blank as needed)
                cells[row, 1].PutValue(0);
            }
        }

        try
        {
            // Ensure the output directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
