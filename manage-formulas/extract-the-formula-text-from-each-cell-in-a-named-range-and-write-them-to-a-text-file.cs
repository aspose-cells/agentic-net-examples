using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace ExtractFormulasFromNamedRange
{
    class Program
    {
        static void Main()
        {
            // Paths
            string workbookPath = "input.xlsx";
            string outputPath = "formulas.txt";

            // Name of the defined range to process
            string rangeName = "MyRange";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                    return;
                }

                // Load the workbook (lifecycle: load)
                Workbook workbook = new Workbook(workbookPath);

                // Retrieve the named range object
                Name namedRange = workbook.Worksheets.Names[rangeName];
                if (namedRange == null)
                {
                    Console.WriteLine($"Error: Named range '{rangeName}' not found.");
                    return;
                }

                // Get all actual ranges referred by the name (could be multiple)
                AsposeRange[] ranges = namedRange.GetRanges();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write formulas to the text file
                using (StreamWriter writer = new StreamWriter(outputPath))
                {
                    foreach (AsposeRange range in ranges)
                    {
                        foreach (Cell cell in range)
                        {
                            if (cell.IsFormula)
                            {
                                writer.WriteLine($"{cell.Name}: {cell.Formula}");
                            }
                        }
                    }
                }

                // (Optional) Save the workbook if any modifications were made (lifecycle: save)
                workbook.Save("output.xlsx");
                Console.WriteLine("Formula extraction completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}