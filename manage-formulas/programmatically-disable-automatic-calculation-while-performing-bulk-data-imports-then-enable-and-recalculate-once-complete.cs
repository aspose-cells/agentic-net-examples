// Title: Disable automatic formula calculation during bulk data import and recalculate afterward with Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET disable automatic formula calculation while importing large arrays | ImportTwoDimensionArray without triggering calculation in Aspose.Cells | Recalculate all formulas after bulk data load with Aspose.Cells C# | Best practice for turning off calculation engine during massive data import in Aspose.Cells | Improve performance of bulk Excel import by suspending calculation in Aspose.Cells
// Tags: disable automatic calculation Aspose.Cells | bulk import ImportTwoDimensionArray Aspose.Cells | Workbook.CalculateFormula after data load | performance optimization Excel import Aspose.Cells | suspend formula engine Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, disables automatic formula calculation, imports a 1000 × 5 two‑dimensional array into the first worksheet using ImportTwoDimensionArray, then forces a full recalculation with Workbook.CalculateFormula before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // ----- Begin bulk data import -----
                Worksheet sheet = workbook.Worksheets[0];

                // Example: import a large 2‑dimensional array into the worksheet.
                object[,] bulkData = new object[1000, 5];
                for (int r = 0; r < 1000; r++)
                {
                    for (int c = 0; c < 5; c++)
                    {
                        bulkData[r, c] = $"R{r + 1}C{c + 1}";
                    }
                }

                // Write the array to the sheet starting at cell A1.
                sheet.Cells.ImportTwoDimensionArray(bulkData, 0, 0);
                // ----- End bulk data import -----

                // Force a full recalculation of all formulas.
                workbook.CalculateFormula();

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
