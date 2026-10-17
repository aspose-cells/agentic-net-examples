// Title: Convert string cells containing French‑formatted numbers to numeric values in a defined Excel range using Aspose.Cells for .NET
// AI Prompts: Load an existing workbook, iterate through cells A1:C10, parse each string with CultureInfo("fr-FR") as a double, and replace the cell value using Aspose.Cells. | Transform all string cells that represent numbers with a comma decimal separator into numeric cell types within a specified range, then save the workbook.
// Common Searches: Aspose.Cells how to convert comma‑separated numeric strings to numbers in a range | C# parse French formatted numbers in Excel cells using Aspose.Cells | Replace string values with doubles in specific cells with Aspose.Cells .NET | Convert string cells to numeric type based on CultureInfo in Aspose.Cells | Save workbook after changing cell types from string to number Aspose.Cells
// Tags: parse culture‑specific numeric strings Aspose.Cells | replace string cell values with numbers Aspose.Cells | range cell type conversion Excel .NET | French decimal separator parsing Aspose.Cells | save workbook after cell value conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// // Loads input.xlsx, iterates over cells A1:C10, parses each string using French culture (comma as decimal separator), replaces the string with a double value, and saves the result to output.xlsx using Aspose.Cells.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load an existing workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the address of the range that contains numeric strings
            string rangeAddress = "A1:C10";

            // Create a Range object for the specified address (use fully qualified name to avoid ambiguity)
            Aspose.Cells.Range range = sheet.Cells.CreateRange(rangeAddress);

            // CultureInfo that defines the decimal separator (e.g., French uses comma)
            CultureInfo culture = new CultureInfo("fr-FR");

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                // Process only cells that currently hold string values
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue.Trim();

                    // Attempt to parse the string using the specified culture
                    if (double.TryParse(text, NumberStyles.Any, culture, out double numericValue))
                    {
                        // Replace the string with the parsed numeric value
                        cell.PutValue(numericValue);
                    }
                }
            }

            // Save the modified workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
