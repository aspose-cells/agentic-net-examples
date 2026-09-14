// Title: Parse formatted currency strings in a column and write numeric values to the adjacent column using Aspose.Cells for .NET
// AI Prompts: Read each cell in column A, strip currency symbols, convert the displayed text to a double, and store the result in column B with Aspose.Cells. | Implement a reusable method that cleans a formatted currency string and parses it using invariant and current culture settings in C#. | Enhance the example to support additional symbols and custom number formats while preserving the original workbook.
// Common Searches: asp.net parse currency formatted cell to numeric value using Aspose.Cells | c# extract numeric amount from Excel cell showing $ or € with Aspose.Cells | how to write parsed double to adjacent column in an Aspose.Cells workbook | culture-aware currency parsing example for Aspose.Cells .NET
// Tags: currency string to double Aspose.Cells | parse formatted cell value C# | write numeric result to adjacent column Aspose.Cells | culture-aware number parsing Aspose.Cells | remove currency symbols Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

namespace CurrencyParsingExample
{
    // The program loads an Excel workbook, iterates through each row in column A, retrieves the cell's displayed string, removes common currency symbols and whitespace, attempts to parse the cleaned text to a double using invariant and current culture formats, writes the numeric value to the neighboring cell in column B, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the range to process (e.g., column A)
            int maxRow = cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                // Get the cell in column A
                Cell sourceCell = cells[row, 0];

                // Retrieve the formatted string value (as seen in Excel)
                string formattedText = sourceCell.StringValue;

                // Attempt to parse currency values (e.g., "$1,234.56", "€1.234,56")
                // Remove common currency symbols and whitespace
                string cleaned = formattedText.Replace("$", "")
                                             .Replace("€", "")
                                             .Replace("£", "")
                                             .Replace("¥", "")
                                             .Replace(" ", "")
                                             .Trim();

                // Remove grouping separators (commas or periods depending on culture)
                // First try invariant culture (comma as thousands, dot as decimal)
                double numericValue;
                bool parsed = double.TryParse(cleaned,
                                             NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign,
                                             CultureInfo.InvariantCulture,
                                             out numericValue);

                // If invariant parsing fails, try current culture (handles cases like "1.234,56")
                if (!parsed)
                {
                    parsed = double.TryParse(cleaned,
                                            NumberStyles.Any,
                                            CultureInfo.CurrentCulture,
                                            out numericValue);
                }

                // If parsing succeeded, store the numeric value in the adjacent cell (column B)
                if (parsed)
                {
                    Cell targetCell = cells[row, 1];
                    targetCell.PutValue(numericValue);
                }
                else
                {
                    // Optionally, you can leave the target cell empty or write an indicator
                    // cells[row, 1].PutValue("N/A");
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
