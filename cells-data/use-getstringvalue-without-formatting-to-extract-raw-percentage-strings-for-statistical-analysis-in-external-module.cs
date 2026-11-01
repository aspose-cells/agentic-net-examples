// Title: Read raw percentage strings from specific Excel cells using Aspose.Cells for .NET and calculate their average
// AI Prompts: Load an .xlsx workbook with Aspose.Cells and retrieve the exact text stored in cells B2, B3, and B4 without applying any number format. | Convert the retrieved percentage strings to double values using CultureInfo.InvariantCulture and compute their arithmetic mean. | Print the original raw strings and the calculated average as invariant‑culture strings to the console.
// Common Searches: Aspose.Cells get cell value as raw text ignoring number format in C# | How to read percentage values from Excel as plain strings with Aspose.Cells .NET | Calculate average of raw Excel percentage strings using Aspose.Cells in a console application
// Tags: Aspose.Cells read raw cell string | extract unformatted percentage value Excel .NET | parse raw Excel string to double invariant culture | compute average of Excel cell values C# | Aspose.Cells workbook cell value retrieval

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' with Aspose.Cells, reads the stored text of cells B2, B3, and B4 using the Value property converted to string, parses these raw percentage strings to double with InvariantCulture, calculates their average, and prints both the original raw strings and the average as invariant‑culture strings.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Retrieve raw cell values (as stored in the file) for B2, B3, B4
            string rawB2 = worksheet.Cells["B2"].Value?.ToString() ?? string.Empty;
            string rawB3 = worksheet.Cells["B3"].Value?.ToString() ?? string.Empty;
            string rawB4 = worksheet.Cells["B4"].Value?.ToString() ?? string.Empty;

            // Output the raw strings for statistical analysis
            Console.WriteLine($"Raw B2: {rawB2}");
            Console.WriteLine($"Raw B3: {rawB3}");
            Console.WriteLine($"Raw B4: {rawB4}");

            // Convert raw strings to numeric values for analysis
            double valB2 = double.Parse(rawB2, CultureInfo.InvariantCulture);
            double valB3 = double.Parse(rawB3, CultureInfo.InvariantCulture);
            double valB4 = double.Parse(rawB4, CultureInfo.InvariantCulture);

            // Calculate average
            double average = (valB2 + valB3 + valB4) / 3.0;
            Console.WriteLine($"Average (as decimal): {average}");

            // Average as a raw string (e.g., "0.4233333333")
            string averageRawString = average.ToString(CultureInfo.InvariantCulture);
            Console.WriteLine($"Average raw string: {averageRawString}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
