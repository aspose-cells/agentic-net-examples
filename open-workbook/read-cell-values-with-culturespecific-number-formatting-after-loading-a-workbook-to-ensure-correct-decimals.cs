// Title: Read Excel cell values with German (de‑DE) number formatting using Aspose.Cells for .NET
// Description: Shows how to load an XLSX workbook with Aspose.Cells LoadOptions that set CultureInfo to de‑DE, so numbers using a comma as decimal separator are parsed correctly. The example extracts the numeric value via DoubleValue or by parsing a string cell and also returns the formatted string representation.
// Keywords: Aspose.Cells | LoadOptions | CultureInfo | de-DE | German number format | comma decimal separator | read numeric cell C# | .NET Excel parsing | culture‑aware Excel reading | Excel workbook loading
// Common Searches: Aspose.Cells load workbook with German locale | read Excel numbers with comma decimal separator in C# | set CultureInfo for Excel file using Aspose.Cells | parse string cell as double with specific culture Aspose | how to handle locale‑specific number formats in Aspose.Cells
// Developer Intent: Load an Excel workbook with a defined locale and retrieve numeric values accurately, respecting the locale's decimal separator.
// Use Cases: Import a German‑formatted spreadsheet and obtain precise double values from cells. | Convert string cells that contain numbers with commas into numeric types using the workbook’s CultureInfo. | Display both the raw double and the locale‑formatted string as they appear in the worksheet.
// AI Prompts: Generate code to read Excel numbers using French (fr‑FR) culture with Aspose.Cells. | Provide an example that writes a double back to a worksheet while preserving the original comma decimal format. | Explain how to detect the workbook’s default CultureInfo after loading and apply it to custom number formatting.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCultureReadDemo
{
    // Shows how to load an XLSX workbook with Aspose.Cells LoadOptions that set CultureInfo to de‑DE, so numbers using a comma as decimal separator are parsed correctly. The example extracts the numeric value via DoubleValue or by parsing a string cell and also returns the formatted string representation.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file that contains numbers formatted with a comma as decimal separator
            string inputFile = "sample.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                return;
            }

            try
            {
                // Create LoadOptions and set the culture to German (de-DE) which uses ',' as decimal separator
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    CultureInfo = new CultureInfo("de-DE")
                };

                // Load the workbook using the specified LoadOptions
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Access the first worksheet and a cell that contains a numeric value
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"]; // Adjust the address as needed

                double numericValue = double.NaN;
                string formattedValue = string.Empty;

                // Retrieve the numeric value in a culture‑aware way
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericValue = cell.DoubleValue;
                }
                else if (cell.Type == CellValueType.IsString)
                {
                    // Attempt to parse the string using the specified culture
                    string raw = cell.StringValue;
                    if (double.TryParse(raw, NumberStyles.Any, loadOptions.CultureInfo, out double parsed))
                    {
                        numericValue = parsed;
                    }
                }

                // Retrieve the formatted string representation (uses the workbook's culture settings)
                formattedValue = cell.StringValue;

                // Output the results
                Console.WriteLine($"Numeric value (DoubleValue): {numericValue}");
                Console.WriteLine($"Formatted string (StringValue): {formattedValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
