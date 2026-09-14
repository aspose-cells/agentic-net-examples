// Title: Load CSV with high‑precision decimals using Aspose.Cells TxtLoadOptions and a custom ICustomParser in C#
// AI Prompts: Implement an ICustomParser that converts string values to decimal and assign it to TxtLoadOptions.PreferredParsers for the numeric column when loading a CSV file. | Enable TxtLoadOptions.KeepPrecision, load CSV data from a MemoryStream into a Workbook, verify cell types and values, and save the workbook as an XLSX file.
// Common Searches: Aspose.Cells how to use TxtLoadOptions.PreferredParsers for custom numeric parsing in C# | preserve decimal precision when importing CSV with Aspose.Cells | example of custom ICustomParser for decimal values in Aspose.Cells CSV import | load CSV with high precision numbers and keep precision property Aspose.Cells | C# load CSV to Excel with column‑specific parser using Aspose.Cells
// Tags: custom ICustomParser decimal parsing Aspose.Cells | TxtLoadOptions PreferredParsers CSV import | KeepPrecision property CSV to Excel | high precision numeric values CSV Aspose.Cells | load CSV with column‑specific parser C#

using System;
using System.IO;
using System.Text;
using System.Globalization;
using Aspose.Cells;

// The example defines a DecimalParser that implements ICustomParser to parse strings as decimal, assigns it to TxtLoadOptions.PreferredParsers (null for the first column, DecimalParser for the second), enables KeepPrecision, loads CSV data from a MemoryStream into a Workbook, prints cell types and values to confirm precise parsing, and saves the result as PreciseNumbers.xlsx.
class Program
{
    // Custom parser that tries to parse a value as decimal to retain full precision.
    private class DecimalParser : ICustomParser
    {
        public bool Parse(string value, out object result)
        {
            if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal dec))
            {
                result = dec;
                return true;
            }
            result = null;
            return false;
        }

        public object ParseObject(string value)
        {
            if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal dec))
                return dec;
            return value;
        }

        public string GetFormat()
        {
            return "Decimal";
        }
    }

    static void Main()
    {
        // Sample CSV containing high‑precision numeric values.
        string csvData = "ID,Value\n1,1234567890.123456789\n2,0.00000000123456789";

        // Create TxtLoadOptions for CSV loading.
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        // Assign custom parsers: first column uses default parser (null), second column uses DecimalParser.
        loadOptions.PreferredParsers = new ICustomParser[] { null, new DecimalParser() };
        // Keep precision for long string values (optional but demonstrates the property).
        loadOptions.KeepPrecision = true;

        // Load the CSV data into a workbook using the configured options.
        using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(ms, loadOptions);
            Cells cells = workbook.Worksheets[0].Cells;

            // Output cell types and values to verify parsing.
            Console.WriteLine($"A2 Type: {cells[1, 0].Type}, Value: {cells[1, 0].StringValue}");
            Console.WriteLine($"B2 Type: {cells[1, 1].Type}, Value: {cells[1, 1].Value}");

            // Save the workbook to an Excel file.
            workbook.Save("PreciseNumbers.xlsx", SaveFormat.Xlsx);
        }
    }
}
