using System;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCustomDateParserDemo
{
    // Custom parser that interprets dates in "dd/MM/yyyy" format.
    public class DateParser : ICustomParser
    {
        private string _lastFormat;

        // Parses the string value. Returns a DateTime object if the format matches,
        // otherwise returns the original string.
        public object ParseObject(string value)
        {
            if (DateTime.TryParseExact(value,
                                       "dd/MM/yyyy",
                                       CultureInfo.InvariantCulture,
                                       DateTimeStyles.None,
                                       out DateTime result))
            {
                _lastFormat = "dd/MM/yyyy";
                return result;
            }

            // Parsing failed – return the original string.
            _lastFormat = null;
            return value;
        }

        // Returns the format used during the last successful parse.
        public string GetFormat()
        {
            return _lastFormat;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Sample CSV data where the first column contains dates in dd/MM/yyyy format.
                string csvData = "\"15/08/2023\",123.45\n\"01/01/2024\",678.90";

                // Create TxtLoadOptions for CSV and assign the custom date parser.
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                {
                    // Enable automatic conversion of numeric values.
                    ConvertNumericData = true,
                    // Use the custom date parser.
                    PreferredParsers = new ICustomParser[] { new DateParser() }
                };

                // Load the workbook from the CSV data using the custom parser.
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
                {
                    Workbook workbook = new Workbook(ms, loadOptions);
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Demonstrate the parsed results.
                    Console.WriteLine("A1 Type: " + cells[0, 0].Type);               // Expected: DateTime
                    Console.WriteLine("A1 Value: " + cells[0, 0].DateTimeValue);   // Expected: 15/08/2023
                    Console.WriteLine("B1 Type: " + cells[0, 1].Type);               // Expected: Double
                    Console.WriteLine("B1 Value: " + cells[0, 1].DoubleValue);     // Expected: 123.45

                    // Prepare output path.
                    string outputPath = "OutputWithCustomDateParser.xlsx";

                    // Ensure the directory exists before saving.
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the workbook to an Excel file.
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}