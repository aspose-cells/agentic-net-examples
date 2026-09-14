// Title: How to parse dd/MM/yyyy dates from a CSV file using a custom ICustomParser with PreferredParsers in Aspose.Cells for .NET
// AI Prompts: Write a class that implements ICustomParser to transform strings formatted as dd/MM/yyyy into DateTime values. | Configure TxtLoadOptions with ConvertDateTimeData enabled, assign the custom parser to PreferredParsers, load CSV data into a Workbook, verify that the first column cells are DateTime type, and save the workbook as XLSX.
// Common Searches: Aspose.Cells how to parse European date format from CSV | assign a date parser to PreferredParsers for column 0 in TxtLoadOptions | enable automatic date conversion when importing CSV with Aspose.Cells .NET | sample code for loading CSV with custom date handling in C# | convert dd/MM/yyyy strings to DateTime during CSV import Aspose.Cells
// Tags: date parser class implementing ICustomParser Aspose.Cells | PreferredParsers setup for CSV loading .NET | ConvertDateTimeData option in TxtLoadOptions | handling dd/MM/yyyy dates in CSV with Aspose.Cells | CSV to Workbook conversion with custom parser Aspose.Cells

using System;
using System.IO;
using System.Globalization;
using System.Text;
using Aspose.Cells;

// Custom parser that interprets dates in "dd/MM/yyyy" format
// The example defines a CustomDateParser that implements ICustomParser to recognize strings in the dd/MM/yyyy pattern and return DateTime objects. It assigns this parser to the PreferredParsers array of TxtLoadOptions, enables ConvertDateTimeData, loads CSV content from a memory stream into a Workbook, checks that the first column cells are stored as DateTime, and finally saves the workbook as an XLSX file.
class CustomDateParser : ICustomParser
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

class Program
{
    static void Main()
    {
        try
        {
            // Sample CSV data where the first column contains dates in dd/MM/yyyy format
            string csvData = "01/12/2023,100\n15/01/2024,200";

            // Create TxtLoadOptions for CSV loading
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                // Assign the custom date parser to the first column (index 0)
                PreferredParsers = new ICustomParser[] { new CustomDateParser() },
                // Ensure that string values are attempted to be converted to dates
                ConvertDateTimeData = true
            };

            // Load the workbook from the CSV data using the custom parser
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                Workbook workbook = new Workbook(ms, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Demonstrate that the first column is parsed as DateTime
                Console.WriteLine("A1 Type: " + cells[0, 0].Type); // Should be DateTime
                Console.WriteLine("A1 Date: " + cells[0, 0].DateTimeValue.ToString("dd/MM/yyyy"));

                // Second column should be numeric; check type before accessing DoubleValue
                Console.WriteLine("B1 Type: " + cells[0, 1].Type);
                if (cells[0, 1].Type == CellValueType.IsNumeric)
                {
                    Console.WriteLine("B1 Value: " + cells[0, 1].DoubleValue);
                }
                else
                {
                    Console.WriteLine("B1 Value is not numeric. Raw value: " + cells[0, 1].StringValue);
                }

                // Save the resulting workbook safely
                string outputPath = "Output.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Error saving workbook: " + saveEx.Message);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unhandled exception: " + ex.Message);
        }
    }
}
