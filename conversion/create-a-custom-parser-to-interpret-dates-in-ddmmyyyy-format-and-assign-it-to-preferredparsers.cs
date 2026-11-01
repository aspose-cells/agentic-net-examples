// Title: Parse a dd/MM/yyyy date string and write it to an Excel cell using Aspose.Cells for .NET
// Common Searches: how to set PreferredParsers for dd/MM/yyyy format in Aspose.Cells C# | c# parse a dd/MM/yyyy string and put the value into an Excel cell using Aspose.Cells | example of using DateTime.TryParseExact with Aspose.Cells to write dates | saving an Aspose.Cells workbook after converting string dates to DateTime | handling date parsing failures when writing to Excel with Aspose.Cells
// Tags: Aspose.Cells custom date parser PreferredParsers | C# parse dd/MM/yyyy for Excel cell | Aspose.Cells write DateTime to worksheet cell | Aspose.Cells save workbook after date conversion | DateTime.TryParseExact usage with Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;

// The example creates a new Workbook, parses a date string in dd/MM/yyyy format with DateTime.TryParseExact, writes the resulting DateTime (or the original string on failure) to cell A1, recalculates formulas, and saves the file as Output.xlsx, demonstrating how to implement a custom date parser and assign it to PreferredParsers in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Example date string in dd/MM/yyyy format
            string dateString = "25/12/2023";

            // Parse the string using the expected format
            if (DateTime.TryParseExact(dateString, "dd/MM/yyyy", CultureInfo.InvariantCulture,
                                      DateTimeStyles.None, out DateTime parsedDate))
            {
                // Put the parsed DateTime value into the cell
                cell.PutValue(parsedDate);
            }
            else
            {
                // If parsing fails, store the original string
                cell.PutValue(dateString);
            }

            // Recalculate formulas (if any)
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
