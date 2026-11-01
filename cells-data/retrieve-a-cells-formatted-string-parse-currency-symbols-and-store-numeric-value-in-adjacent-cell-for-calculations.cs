// Title: Read formatted currency strings from Excel cells and write parsed numeric values to the adjacent column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells get displayed string of a cell in C# | C# parse currency symbols from Excel formatted values with Aspose.Cells | Convert Excel currency formatted text to decimal and store in next column using Aspose.Cells | Handle parentheses for negative numbers when extracting values from Excel in .NET | Batch process rows to extract numeric values from formatted currency cells with Aspose.Cells
// Tags: GetStringValue display string Aspose.Cells | regex strip currency symbols C# | parse formatted currency to decimal Aspose.Cells | write numeric value to adjacent cell Excel Aspose.Cells | handle parentheses negative numbers Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, iterates through each row of the first worksheet, retrieves each cell's displayed (formatted) string using GetStringValue with the DisplayString strategy, removes currency symbols and whitespace via a regular expression, converts parentheses‑enclosed values to negatives, parses the cleaned text to a decimal using invariant culture, writes the numeric result into the neighboring column, and saves the modified workbook.
class CurrencyParser
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define source (A) and target (B) columns
            int sourceColumn = 0; // A
            int targetColumn = 1; // B

            // Determine the range of rows that contain data
            int startRow = 0;
            int endRow = sheet.Cells.MaxDataRow; // last row with any data

            for (int row = startRow; row <= endRow; row++)
            {
                Cell sourceCell = sheet.Cells[row, sourceColumn];

                // Skip empty or null cells
                if (sourceCell == null || sourceCell.Type == CellValueType.IsNull)
                    continue;

                // Get the displayed (formatted) string value of the cell
                string formattedValue = sourceCell.GetStringValue(CellValueFormatStrategy.DisplayString);

                // Remove currency symbols and whitespace
                string cleaned = Regex.Replace(formattedValue, @"[\p{Sc}\s]", "");

                // Handle parentheses for negative numbers (e.g., (123.45) => -123.45)
                bool isNegative = false;
                if (cleaned.StartsWith("(") && cleaned.EndsWith(")"))
                {
                    isNegative = true;
                    cleaned = cleaned.Trim('(', ')');
                }

                // Parse the numeric part using invariant culture
                if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal numericValue))
                {
                    if (isNegative)
                        numericValue = -numericValue;

                    // Store the numeric value in the adjacent cell
                    Cell targetCell = sheet.Cells[row, targetColumn];
                    targetCell.PutValue(numericValue);
                }
                // If parsing fails, the target cell is left blank
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
