// Title: Create a currency‑formatted report line in C# using Aspose.Cells GetStringValue with WithFormatting option
// Common Searches: how to retrieve a cell value with currency formatting using Aspose.Cells GetStringValue in C# | Aspose.Cells C# get formatted string including currency symbol from Excel cell | generate a console report line with formatted currency value from Excel using Aspose.Cells | C# Aspose.Cells read cell B2 as formatted currency string for reporting | use WithFormatting option to include currency symbols when extracting Excel data with Aspose.Cells
// Tags: GetStringValue WithFormatting Aspose.Cells | formatted currency extraction from Excel C# | Aspose.Cells cell value formatting for reporting | C# console output of Excel currency data | load workbook and read formatted cell Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, accesses cell B2, uses Aspose.Cells GetStringValue with the WithFormatting option (via the StringValue property) to obtain the currency‑formatted text (e.g., $1,234.56), builds a report line such as "Total revenue: $1,234.56", writes it to the console, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Assume the cell B2 contains a numeric value formatted as currency in Excel
            Cell cell = sheet.Cells["B2"];

            // Retrieve the cell's formatted string value (includes currency symbol)
            string formattedValue = cell.StringValue;

            // Build a user‑friendly report line
            string reportLine = $"Total revenue: {formattedValue}";

            // Output the report line
            Console.WriteLine(reportLine);

            // Save the workbook (if any changes were made)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
