using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class DisableLinksUpToDateDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the existing workbook to be loaded
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook from the file system
                Workbook workbook = new Workbook(inputPath);

                // Disable the LinksUpToDate built‑in property to prevent link checks
                workbook.BuiltInDocumentProperties.LinksUpToDate = false;

                // Path where the modified workbook will be saved
                string outputPath = "output.xlsx";

                // Save the workbook with the updated property
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}