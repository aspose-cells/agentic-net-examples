// Title: Parse currency‑formatted cells to numeric values and write them to the adjacent column with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts currency‑styled strings in column A, reads each cell's displayed text via StringValue, strips common symbols and grouping separators, parses the result using invariant and current cultures to handle US and European formats, and writes the numeric value (or the original text on failure) to column B before saving the file.
// Keywords: Aspose.Cells | C# | .NET | currency parsing | formatted cell string | StringValue | CultureInfo | double conversion | Excel financial data | locale aware parsing
// Common Searches: Aspose.Cells read formatted currency string C# | Convert Excel currency text to number with Aspose.Cells | Parse $ and € symbols in Excel cells using .NET | Locale‑aware currency parsing Aspose.Cells example | Write numeric result to adjacent cell in Aspose.Cells
// Developer Intent: Read a cell’s displayed currency text, remove symbols and separators, convert it to a numeric type, and store the result in the next column for calculations.
// Use Cases: Transform mixed US/EU currency strings in a worksheet into pure numbers for aggregation or charting. | Import financial reports where values are stored as formatted text and need to be used in formulas. | Provide a fallback that preserves the original text when parsing fails, ensuring data integrity.
// AI Prompts: Generate Aspose.Cells C# code that extracts a cell's StringValue, removes currency symbols, parses it to double with culture support, and writes the number to the neighboring cell. | Create a reusable method that accepts a Cell object and returns a nullable double after cleaning currency symbols and handling locale‑specific formats. | Suggest improvements to extend the parsing loop for additional symbols, custom formats, and detailed error logging.

using System;
using System.Globalization;
using Aspose.Cells;

// This example creates a workbook, inserts currency‑styled strings in column A, reads each cell's displayed text via StringValue, strips common symbols and grouping separators, parses the result using invariant and current cultures to handle US and European formats, and writes the numeric value (or the original text on failure) to column B before saving the file.
class CurrencyParsingDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Sample data: currency formatted strings in column A
        cells["A1"].PutValue("$1,234.56");
        cells["A2"].PutValue("€2.345,67"); // European format with comma decimal
        cells["A3"].PutValue("£3,210");    // No decimals
        cells["A4"].PutValue("1234");      // Plain number, no currency

        // Optional: apply a built‑in currency number format to column A for visual consistency
        Style currencyStyle = wb.CreateStyle();
        currencyStyle.Number = 164; // Built‑in currency format
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        cells["A1"].SetStyle(currencyStyle, flag);
        cells["A2"].SetStyle(currencyStyle, flag);
        cells["A3"].SetStyle(currencyStyle, flag);
        cells["A4"].SetStyle(currencyStyle, flag);

        // Iterate over used rows in column A
        int maxRow = cells.MaxDataRow;
        for (int row = 0; row <= maxRow; row++)
        {
            Cell srcCell = cells[row, 0];               // Source cell (column A)
            string formattedValue = srcCell.StringValue; // Formatted string including currency symbol

            // Strip common currency symbols and grouping separators
            string cleaned = formattedValue
                .Replace("$", "")
                .Replace("€", "")
                .Replace("£", "")
                .Replace(",", "")
                .Trim();

            // Attempt to parse the cleaned string to a double
            double numericValue;
            bool parsed = double.TryParse(
                cleaned,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out numericValue);

            // If invariant parsing fails, try the current culture (handles cases like "2.345,67")
            if (!parsed)
            {
                parsed = double.TryParse(
                    cleaned,
                    NumberStyles.Any,
                    CultureInfo.CurrentCulture,
                    out numericValue);
            }

            // Write the result to the adjacent cell in column B
            Cell destCell = cells[row, 1];
            if (parsed)
                destCell.PutValue(numericValue); // Store numeric value for calculations
            else
                destCell.PutValue(cleaned);      // Fallback: store the original text
        }

        // Save the workbook (lifecycle rule)
        wb.Save("CurrencyParsingDemo.xlsx");
    }
}
