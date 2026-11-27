// Title: Read raw scientific notation string from an Excel cell using Aspose.Cells GetStringValue in C#
// Common Searches: how to get the original scientific notation text of a cell with Aspose.Cells in C# | using GetStringValue to retrieve unformatted cell values in Aspose.Cells | parse a double from a scientific notation string returned by Aspose.Cells | Aspose.Cells C# example for reading cell value without number formatting | extract raw string from Excel cell that contains 1.23E+04 using Aspose.Cells
// Tags: Aspose.Cells GetStringValue raw string extraction | read scientific notation cell value C# | parse double from unformatted Excel string Aspose | retrieve cell text without formatting Aspose.Cells | handle Excel scientific notation using Aspose.Cells API

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The sample loads an Excel workbook, accesses a target cell, attempts to read its numeric value directly, and if needed falls back to GetStringValue to obtain the raw scientific notation string, trims it, parses it to a double with invariant culture, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Access the cell that contains a number in scientific notation (e.g., "1.23E+04")
            Cell cell = sheet.Cells["A1"];

            double numericValue;

            // Try to obtain the numeric value directly from the cell
            if (cell.Value is double d)
            {
                numericValue = d;
                Console.WriteLine($"Extracted numeric value: {numericValue}");
            }
            else
            {
                // Fallback: get the cell's string representation and parse it
                string rawString = cell.StringValue?.Trim();

                if (double.TryParse(rawString, NumberStyles.Float, CultureInfo.InvariantCulture, out numericValue))
                {
                    Console.WriteLine($"Extracted numeric value from string: {numericValue}");
                }
                else
                {
                    Console.WriteLine($"Unable to parse the string '{rawString}' as a number.");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook if any modifications were made (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
