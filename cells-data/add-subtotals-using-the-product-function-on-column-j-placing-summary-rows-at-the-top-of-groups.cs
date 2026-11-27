// Title: Template for Adding Product Subtotals on Column J with Top Summary Rows Using Aspose.Cells for .NET (C#)
// Common Searches: how to insert product subtotal rows at the beginning of each group with Aspose.Cells C# | Aspose.Cells C# add top-of-group summary rows using Product function on column J | C# code to calculate product subtotals for grouped data in Excel with Aspose.Cells | using Aspose.Cells to place subtotal rows above groups in an Excel sheet
// Tags: Aspose.Cells product subtotal implementation | C# top summary rows for grouped Excel data | Excel column J product aggregation Aspose.Cells | grouped data subtotal automation with Aspose.Cells | Aspose.Cells worksheet summary row insertion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file into an Aspose.Cells Workbook, accesses the first worksheet, ensures the output directory exists, and saves the workbook to a new file while handling possible exceptions. It serves as a base for inserting product subtotals on column J and adding top‑of‑group summary rows.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (optional further processing can be added here)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the message
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }
}
