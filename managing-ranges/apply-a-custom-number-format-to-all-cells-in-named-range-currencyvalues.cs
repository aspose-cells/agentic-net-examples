using System;
using System.IO;
using Aspose.Cells;

namespace ApplyCustomNumberFormat
{
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "CurrencyValues"
                Name namedRange = workbook.Worksheets.Names["CurrencyValues"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'CurrencyValues' not found.");
                    return;
                }

                // Get the A1‑style reference of the named range (e.g., "Sheet1!$B$2:$B$10")
                string rangeRef = namedRange.GetRefersTo(false, false);

                // Create a Range object based on the reference
                // Use fully qualified name to avoid ambiguity with System.Range
                Aspose.Cells.Range range = workbook.Worksheets[0].Cells.CreateRange(rangeRef);

                // Define a custom number format (Euro currency with red negative values)
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "_-€ #,##0.00;[Red]_-€ -#,##0.00";

                // Apply only the number format
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                range.ApplyStyle(customStyle, flag);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}