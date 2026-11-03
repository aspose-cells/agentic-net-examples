// Title: Read a formatted currency cell with Aspose.Cells GetStringValue and convert it to a decimal in C#
// Common Searches: Aspose.Cells C# get cell string value with formatting for currency | How to strip currency symbols from a formatted Excel cell using Aspose.Cells | Convert formatted Excel currency string to decimal in C# with Aspose.Cells | Retrieve formatted value from Excel cell and parse to numeric using Aspose.Cells | C# Aspose.Cells remove $ sign from cell string value
// Tags: GetStringValue with formatting Aspose.Cells | currency string to decimal conversion C# | clean numeric value from formatted Excel cell | Aspose.Cells formatted cell extraction | parse invariant culture numeric string C# | Excel cell value sanitization Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads (or creates) an Excel workbook, reads the formatted currency text from a cell using GetStringValue with formatting, removes non‑numeric characters, parses the cleaned string to a decimal with invariant culture, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Sample data for demonstration
                ws.Cells["B2"].PutValue("$123.45");
                workbook.Save(inputPath);
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the cell that contains the currency value (e.g., B2)
            Cell cell = sheet.Cells["B2"];

            // Retrieve the formatted string (includes currency symbol)
            string formattedValue = cell.StringValue; // formatted representation

            // Remove non‑numeric characters except decimal separators and sign
            string numericString = "";
            foreach (char c in formattedValue)
            {
                if (char.IsDigit(c) || c == '.' || c == ',' || c == '-')
                    numericString += c;
            }

            // Parse the cleaned string to decimal using invariant culture
            decimal amount = decimal.Parse(numericString, NumberStyles.Any, CultureInfo.InvariantCulture);

            // Example usage of the parsed amount
            Console.WriteLine("Parsed amount: " + amount);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
