// Title: Read Excel cell values with culture‑specific number formatting using Aspose.Cells for .NET
// Common Searches: how to read numbers with German decimal separator in Aspose.Cells C# | parse localized numeric string from an Excel cell using Aspose.Cells | Aspose.Cells read cell value as double with specific CultureInfo | handle comma as decimal separator when extracting Excel data with Aspose.Cells | convert Excel cell string to double using culture info in .NET
// Tags: culture-specific number parsing Aspose.Cells | read Excel cell with designated CultureInfo | comma decimal separator processing .NET | fallback to displayed string parsing Aspose.Cells | numeric value extraction from Excel workbook

using System;
using System.Globalization;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, sets a German CultureInfo, reads cell B2, checks whether the stored value is already numeric or a formatted string, parses the string using the specified culture (or falls back to the cell's displayed string), and prints the resulting double using invariant culture.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "input.xlsx";

        // Load the workbook (lifecycle rule)
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the culture that matches the number formatting used in the workbook
        // Example: German culture uses comma as decimal separator
        CultureInfo culture = new CultureInfo("de-DE");

        // Read a specific cell (e.g., B2) that contains a number with culture‑specific formatting
        Cell cell = sheet.Cells["B2"];

        // Retrieve the raw value stored in the cell
        object rawValue = cell.Value;

        double numericValue;

        if (rawValue is double)
        {
            // The cell already holds a numeric value
            numericValue = (double)rawValue;
        }
        else if (rawValue is string)
        {
            // Parse the string using the specified culture
            if (double.TryParse((string)rawValue, NumberStyles.Any, culture, out double parsed))
            {
                numericValue = parsed;
            }
            else
            {
                // As a fallback, use the displayed string (which respects the cell's number format)
                string displayed = cell.StringValue;
                numericValue = double.Parse(displayed, NumberStyles.Any, culture);
            }
        }
        else
        {
            // Handle other possible types (e.g., null) as needed
            numericValue = 0;
        }

        // Output the parsed numeric value (using invariant culture for consistent display)
        Console.WriteLine($"Parsed numeric value: {numericValue.ToString(CultureInfo.InvariantCulture)}");
    }
}
