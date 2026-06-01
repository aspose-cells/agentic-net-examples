using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // NOTE: EnableCalculationChain property may not be available in newer versions.
                // The calculation chain is built automatically when formulas are calculated.

                // Populate source range A1:A5 with numeric values
                for (int i = 0; i < 5; i++)
                {
                    cells[i, 0].PutValue(i + 1); // A1..A5 = 1..5
                }

                // Populate source range B1:B5 with formulas that depend on column A
                for (int i = 0; i < 5; i++)
                {
                    cells[i, 1].Formula = $"A{i + 1}*2"; // B1..B5 = A*2
                }

                // Calculate formulas so that the dependency chain is built
                workbook.CalculateFormula();

                // Define source and destination ranges (use fully qualified Aspose.Cells.Range)
                Aspose.Cells.Range sourceRange = cells.CreateRange("B1:B5");
                Aspose.Cells.Range destRange = cells.CreateRange("D1:D5");

                // Use PasteOptions to copy everything (values, formulas, formats)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                };

                // Perform the copy while preserving calculation dependencies
                destRange.Copy(sourceRange, pasteOptions);

                // Recalculate after copy to ensure dependent cells are updated
                workbook.CalculateFormula();

                // Optional: display results in console for verification
                Console.WriteLine("Source formulas and values:");
                for (int i = 0; i < 5; i++)
                {
                    Console.WriteLine($"B{i + 1}: Formula={cells[i, 1].Formula}, Value={cells[i, 1].Value}");
                }

                Console.WriteLine("\nCopied formulas and values:");
                for (int i = 0; i < 5; i++)
                {
                    Console.WriteLine($"D{i + 1}: Formula={cells[i, 3].Formula}, Value={cells[i, 3].Value}");
                }

                // Save the workbook
                string outputPath = "RangeCopyWithDependencies.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}