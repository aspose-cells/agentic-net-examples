// Title: Custom ICustomParser for dd/MM/yyyy Dates in Aspose.Cells CSV LoadOptions
// Description: Shows how to implement a DdMMyyyyParser that implements ICustomParser, assign it to the second element of TxtLoadOptions.PreferredParsers, and automatically convert a CSV date column (dd/MM/yyyy) into DateTime cells in an Aspose.Cells workbook.
// Keywords: Aspose.Cells | C# custom parser | ICustomParser | PreferredParsers | TxtLoadOptions CSV | dd/MM/yyyy date parsing | ConvertDateTimeData | CSV to Excel conversion
// Common Searches: Aspose.Cells custom date parser example | How to use PreferredParsers with TxtLoadOptions | Parse dd/MM/yyyy dates from CSV in C# | ICustomParser implementation Aspose.Cells | Convert CSV date strings to DateTime cells
// Developer Intent: Automatically transform dd/MM/yyyy date strings in a specific CSV column into native DateTime cells when loading the file with Aspose.Cells.
// Use Cases: Load a CSV where the second column contains dates in dd/MM/yyyy format and store them as proper DateTime values in the workbook. | Preserve original text for cells that do not match the expected date pattern. | Export the parsed workbook to Excel while keeping correct date formatting for downstream processing.
// AI Prompts: Create an ICustomParser in C# that parses "dd/MM/yyyy" dates and integrates it with TxtLoadOptions.PreferredParsers for CSV loading in Aspose.Cells. | Show how to fallback to the original string when a custom parser fails to parse a value. | Demonstrate retrieving the last successful format from a custom ICustomParser after loading a CSV workbook.

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

// Shows how to implement a DdMMyyyyParser that implements ICustomParser, assign it to the second element of TxtLoadOptions.PreferredParsers, and automatically convert a CSV date column (dd/MM/yyyy) into DateTime cells in an Aspose.Cells workbook.
class Program
{
    // Custom parser that interprets dates in "dd/MM/yyyy" format
    private class DdMMyyyyParser : ICustomParser
    {
        private string _lastFormat;

        // Parse the string; if it matches the expected date format, return DateTime
        public object ParseObject(string value)
        {
            if (DateTime.TryParseExact(value, "dd/MM/yyyy", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime dt))
            {
                _lastFormat = "dd/MM/yyyy";
                return dt;
            }

            // If parsing fails, return the original string (default handling)
            _lastFormat = null;
            return value;
        }

        // Return the format used during the last successful parse
        public string GetFormat()
        {
            return _lastFormat;
        }
    }

    static void Main()
    {
        // Sample CSV data where the second column contains dates in dd/MM/yyyy format
        string csvData = "Name,Date\nJohn,25/12/2023\nDoe,01/01/2024";

        // Create TxtLoadOptions for CSV and assign the custom parser to the second column
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        // parsers[0] -> first column (no custom parser), parsers[1] -> second column (our date parser)
        loadOptions.PreferredParsers = new ICustomParser[] { null, new DdMMyyyyParser() };
        loadOptions.ConvertDateTimeData = true; // Ensure date strings are converted to DateTime

        // Load the workbook from the CSV data using the defined options
        using (MemoryStream ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(ms, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Demonstrate that the date column was parsed as DateTime
            Console.WriteLine("A2 (Name): " + cells["A2"].StringValue);
            Console.WriteLine("B2 (Date) Type: " + cells["B2"].Type);
            Console.WriteLine("B2 (Date) Value: " + cells["B2"].DateTimeValue.ToString("dd/MM/yyyy"));

            // Save the workbook (optional)
            workbook.Save("ParsedDates.xlsx");
        }
    }
}
