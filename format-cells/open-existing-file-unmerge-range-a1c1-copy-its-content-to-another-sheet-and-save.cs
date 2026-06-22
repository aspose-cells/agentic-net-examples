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
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Load existing workbook if it exists; otherwise create a new workbook
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Source worksheet (first sheet)
                Worksheet srcSheet = workbook.Worksheets[0];

                // Unmerge the range A1:C1 (zero‑based indices: row 0, column 0, 1 row, 3 columns)
                srcSheet.Cells.UnMerge(0, 0, 1, 3);

                // Add a new worksheet to receive the copied content
                Worksheet destSheet = workbook.Worksheets.Add("CopySheet");

                // Define source and destination ranges using Aspose.Cells.Range to avoid ambiguity with System.Range
                Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange("A1:C1");
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:C1");

                // Set paste options to copy all content (values, formats, etc.)
                PasteOptions options = new PasteOptions
                {
                    PasteType = PasteType.All,
                    SkipBlanks = false,
                    Transpose = false
                };

                // Copy the source range to the destination range
                srcRange.Copy(destRange, options);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}