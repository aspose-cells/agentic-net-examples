// Title: Use TxtLoadOptions.PreferredParsers with a custom numeric parser to load CSV and enforce decimal precision in Aspose.Cells for .NET
// Common Searches: how to set a custom numeric parser for CSV import using TxtLoadOptions in Aspise.Cells C# | Aspose.Cells load CSV with specific decimal rounding using ICustomParser | C# example of enforcing 6 decimal places when reading CSV with Aspose.Cells | configure TxtLoadOptions.PreferredParsers to round numbers during CSV to workbook conversion | use CustomNumericParser with Aspose.Cells TxtLoadOptions to control precision
// Tags: ICustomParser implementation for rounding numbers | TxtLoadOptions PreferredParsers for CSV files | decimal precision control on CSV load | Aspose.Cells workbook load with precision settings | CSV to XLSX conversion preserving numeric precision

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Custom numeric parser that parses numbers with high precision and rounds to a specified number of decimal places
// The sample defines a CustomNumericParser that implements ICustomParser to round parsed decimal values to a configurable number of places (e.g., 6). It configures TxtLoadOptions with a comma separator and assigns the custom parser via the PreferredParsers array, then loads a CSV file into a Workbook and saves it as XLSX, ensuring numeric precision is maintained while handling missing files and runtime errors.
public class CustomNumericParser : ICustomParser
{
    private readonly int _decimalPlaces;

    public CustomNumericParser(int decimalPlaces = 10)
    {
        _decimalPlaces = decimalPlaces;
    }

    // Parse a string value and return a decimal rounded to the required precision, or the original string if parsing fails
    public object ParseObject(string value)
    {
        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal number))
        {
            number = Math.Round(number, _decimalPlaces, MidpointRounding.AwayFromZero);
            return number;
        }

        return value;
    }

    // No specific format required
    public string GetFormat()
    {
        return null;
    }

    // Optional helper matching older signature (not required by the interface)
    public object Parse(string value, Cell cell)
    {
        return ParseObject(value);
    }
}

public class CsvLoader
{
    public static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Input file not found: {csvPath}");
            return;
        }

        // Configure TxtLoadOptions with the custom numeric parser
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            // Ensure the file is treated as a CSV (comma‑separated values)
            Separator = ',',
            // Assign the custom parser via the PreferredParsers array
            PreferredParsers = new ICustomParser[] { new CustomNumericParser(decimalPlaces: 6) }
        };

        try
        {
            // Load the CSV file using the configured options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Save the workbook to another format to verify the precision enforcement
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
