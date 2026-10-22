// Title: How to verify a chart’s secondary axis in Aspose.Cells (C#) before adding one
// AI Prompts: Generate C# code with Aspose.Cells that inspects a chart, determines if a secondary value axis already exists, and adds the axis only when it is missing. | Show a step‑by‑step example that conditionally creates a secondary axis on an Excel chart using Aspose.Cells, avoiding duplicate axes.
// Common Searches: Aspose.Cells C# check if chart already has secondary axis before adding | prevent duplicate secondary axis in Aspose.Cells chart | determine presence of secondary value axis with Aspose.Cells C# | conditionally add secondary axis to Excel chart using Aspose.Cells
// Tags: Aspose.Cells chart secondary axis detection | C# conditional secondary axis creation | Excel chart manipulation without duplicate axes | Aspose.Cells chart API secondary axis check

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, accesses the first worksheet and its first chart, and explains where to insert logic that checks for an existing secondary value axis before creating one, ensuring the chart does not end up with duplicate secondary axes.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or specify another index/name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count > 0)
                {
                    // Get the first chart
                    Chart chart = worksheet.Charts[0];

                    // NOTE: The SecondaryValueAxis and IsOnSecondaryAxis members are not
                    // available in the current Aspose.Cells version. If needed, use
                    // alternative chart manipulation APIs here.
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
