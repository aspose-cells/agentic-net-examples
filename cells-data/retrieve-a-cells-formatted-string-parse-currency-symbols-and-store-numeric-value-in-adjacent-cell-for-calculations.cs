using System;
using Aspose.Cells;

namespace CurrencyParsingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Sample data: put some currency formatted strings in column A
            // ------------------------------------------------------------
            // Apply a custom currency format to the cells so that StringValue
            // returns the formatted representation (e.g., "$1,234.56").
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";

            // Row 0
            cells["A1"].PutValue(1234.56);
            cells["A1"].SetStyle(currencyStyle);

            // Row 1
            cells["A2"].PutValue(987.0);
            cells["A2"].SetStyle(currencyStyle);

            // Row 2 – a value that is not a currency (will stay as is)
            cells["A3"].PutValue("Not a number");

            // ------------------------------------------------------------
            // Process each used row in column A
            // ------------------------------------------------------------
            int maxRow = cells.MaxDataRow; // last row that contains data
            for (int row = 0; row <= maxRow; row++)
            {
                Cell sourceCell = cells[row, 0]; // column A
                // Get the formatted string as it appears in Excel
                string formattedText = sourceCell.StringValue;

                // Try to convert the formatted text to a numeric value using
                // the NUMBERVALUE worksheet function. This handles currency
                // symbols, thousand separators, etc.
                object numericResult = null;
                try
                {
                    // The formula is =NUMBERVALUE("formattedText")
                    // Escape any double quotes inside the text.
                    string escapedText = formattedText.Replace("\"", "\"\"");
                    string formula = $"=NUMBERVALUE(\"{escapedText}\")";
                    numericResult = sheet.CalculateFormula(formula);
                }
                catch
                {
                    // If conversion fails, leave numericResult as null
                }

                // If conversion succeeded and the result is a double, store it
                // in the adjacent cell (column B) for further calculations.
                Cell targetCell = cells[row, 1]; // column B
                if (numericResult is double d)
                {
                    targetCell.PutValue(d);
                }
                else
                {
                    // Preserve the original text if it cannot be parsed
                    targetCell.PutValue(formattedText);
                }
            }

            // ------------------------------------------------------------
            // Save the workbook to verify the result
            // ------------------------------------------------------------
            workbook.Save("CurrencyParsingResult.xlsx");
        }
    }
}