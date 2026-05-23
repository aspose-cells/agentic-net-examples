using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "ProtectedWorkbook.xlsx";
                const string outputPath = "FormattedCopy.xlsx";

                // Ensure the source workbook exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the password‑protected workbook
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = "myPassword"
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define source and destination ranges (use fully qualified type to avoid ambiguity)
                Aspose.Cells.Range sourceRange = cells.CreateRange("O1:O5");
                Aspose.Cells.Range destinationRange = cells.CreateRange("P1:P5");

                // Copy only the formatting from source to destination
                destinationRange.CopyStyle(sourceRange);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}