// Title: Extract raw scientific‑notation text from an Excel cell using Aspose.Cells GetStringValue and convert it to a double in C#
// AI Prompts: Use Aspose.Cells GetStringValue to read the exact text of a cell (e.g., A1) without applying any number format, then parse the resulting scientific‑notation string to a double with InvariantCulture in C#. | Show how to safely handle cases where GetStringValue returns a non‑numeric string and convert only valid scientific‑notation values to double using TryParse in a .NET application.
// Common Searches: Aspose.Cells GetStringValue returns scientific notation string C# | How to read an Excel cell value as raw text without formatting using Aspose.Cells | Parse scientific notation from Excel cell to double in .NET | GetStringValue vs Value property Aspose.Cells example | Convert raw Excel cell string to double with invariant culture C#
// Tags: raw cell text extraction GetStringValue | scientific notation parsing double C# | read unformatted Excel cell value .NET | convert raw Excel string to numeric Aspose.Cells | handle non‑numeric cell content C#

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads an input.xlsx workbook, accesses cell A1 on the first worksheet, retrieves the cell's raw string using GetStringValue (bypassing any number formatting), attempts to parse that string as a double with invariant culture, reports success or failure, and finally saves the workbook to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Read cell A1
        Cell cell = sheet.Cells["A1"];

        // Retrieve the cell's raw value as a string (without applying any number format)
        string rawValue = cell.Value?.ToString() ?? string.Empty;

        // Attempt to parse the raw string as a double using invariant culture
        if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue))
        {
            Console.WriteLine($"Parsed numeric value: {numericValue}");
        }
        else
        {
            Console.WriteLine($"Failed to parse '{rawValue}' as a number.");
        }

        // Save the workbook (if any modifications were made)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
