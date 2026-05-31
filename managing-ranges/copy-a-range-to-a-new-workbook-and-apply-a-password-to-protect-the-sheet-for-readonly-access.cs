using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyAndProtect
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Define the source range to copy (e.g., A1:C5)
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

                // Create a new workbook for the copied range
                Workbook destWorkbook = new Workbook(); // empty workbook
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define the destination range where the source data will be pasted
                AsposeRange destRange = destSheet.Cells.CreateRange("A1:C5");

                // Copy the source range to the destination range
                destRange.Copy(sourceRange);

                // Protect the destination worksheet with a password for read‑only access
                destSheet.Protect(ProtectionType.All, "readonly", null);

                // Save the new workbook
                destWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}