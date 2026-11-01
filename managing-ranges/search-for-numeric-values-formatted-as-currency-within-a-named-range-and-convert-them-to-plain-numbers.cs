// Title: Convert currency‑formatted numeric cells to General numbers within a named range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change currency formatted cells to General number in a specific named range | remove currency number format from cells in a named range with Aspose.Cells for .NET | detect built‑in and custom currency formats in Excel using Aspose.Cells C# | how to clear custom number format symbols like $ € £ from a named range with Aspose.Cells | iterate over cells in a named range and reset number format to General using Aspose.Cells
// Tags: currency format detection Aspose.Cells | named range number format conversion C# | reset cell style to General Aspose.Cells | custom number format removal Excel C# | iterate named range cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace CurrencyConversionExample
{
    // The example loads an Excel workbook, retrieves a named range, iterates through each cell, identifies numeric cells formatted as currency (both built‑in format 164 and custom formats containing common currency symbols), and updates their style to the General number format (built‑in 0) while clearing any custom format, then saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range (replace "MyNamedRange" with your actual named range)
                AsposeRange namedRange = workbook.Worksheets.GetRangeByName("MyNamedRange");
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyNamedRange' not found.");
                    return;
                }

                // Iterate through each cell in the named range
                foreach (Cell cell in namedRange)
                {
                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        // Retrieve the cell's style
                        Style style = cell.GetStyle();

                        // Determine if the cell is formatted as currency
                        bool isCurrency = false;

                        // Check built‑in number format for currency (164 is a typical built‑in currency format)
                        if (style.Number == 164)
                            isCurrency = true;

                        // Check custom number format for currency symbols
                        if (!isCurrency && !string.IsNullOrEmpty(style.Custom))
                        {
                            string fmt = style.Custom;
                            if (fmt.Contains("$") || fmt.Contains("€") || fmt.Contains("£") || fmt.Contains("¥"))
                                isCurrency = true;
                        }

                        // If the cell is formatted as currency, convert it to a plain number format
                        if (isCurrency)
                        {
                            // Set the number format to General (built‑in format 0)
                            style.Number = 0;               // General format
                            style.Custom = string.Empty;    // Clear any custom format

                            // Apply the updated style back to the cell
                            cell.SetStyle(style);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
