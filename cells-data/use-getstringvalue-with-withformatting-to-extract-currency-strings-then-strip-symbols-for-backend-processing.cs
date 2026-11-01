// Title: Extract a formatted currency string from an Excel cell using Aspose.Cells for .NET and convert it to a decimal
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, reads a cell's formatted currency text (including the symbol), strips non‑numeric characters, and parses it into a decimal. | Show how to use Cell.GetStringValue with the WithFormatting option to obtain a currency string and then clean it for backend processing. | Provide a complete example that reads cell B2, removes currency symbols and extra characters, handles different decimal separators, and saves the workbook after conversion.
// Common Searches: Aspose.Cells C# read cell value with formatting and remove currency symbol | How to parse formatted currency from Excel using Aspose.Cells .NET | Convert Excel formatted amount to decimal in C# with Aspose.Cells | GetStringValue WithFormatting currency example Aspose.Cells | Strip non‑numeric characters from Excel cell string in C#
// Tags: read formatted currency Aspose.Cells C# | parse cell string to decimal Aspose.Cells | strip currency symbols from Excel value | Cell.GetStringValue WithFormatting example | backend numeric conversion Excel C#

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;
using System.Text;

// The sample loads an input.xlsx workbook, retrieves the formatted currency string from cell B2 using Aspose.Cells (including the currency symbol), removes all non‑numeric characters while preserving decimal separators and sign, parses the cleaned string into a decimal, outputs the amount, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Get the formatted string value of the cell (includes currency symbol)
            Cell cell = sheet.Cells["B2"];
            string formattedValue = cell.StringValue;

            // Remove non‑numeric characters except decimal separators and sign
            StringBuilder sb = new StringBuilder();
            foreach (char c in formattedValue)
            {
                if (char.IsDigit(c) || c == '.' || c == ',' || c == '-')
                    sb.Append(c);
            }
            string numericString = sb.ToString();

            // Convert the cleaned string to a decimal
            if (decimal.TryParse(numericString, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amount))
            {
                Console.WriteLine("Parsed amount: " + amount);
            }
            else
            {
                Console.WriteLine("Failed to parse amount from: " + formattedValue);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
