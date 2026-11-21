// Title: Calculate the median of a range and store it in another cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to use MEDIAN worksheet function on a range | Aspose.Cells calculate median of A1:A10 and write result to B1 | Get numeric result from CalculateFormula in Aspose.Cells C# | Save workbook after computing statistical functions with Aspose.Cells | Populate cells with sequential numbers using Aspose.Cells C# example
// Tags: median worksheet function aspose.cells c# | calculateformula numeric result aspose.cells | populate sequential values aspose.cells worksheet | write double value to cell aspose.cells | save workbook as xlsx aspose.cells .net

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMedianExample
{
    // The example creates a new Workbook, fills cells A1‑A10 with values 1‑10, uses the MEDIAN worksheet function via CalculateFormula to obtain the median, converts the result to a double, writes the median into cell B1, ensures the output directory exists, and saves the file as MedianExample.xlsx while handling exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in A1:A10 (values 1 through 10).
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue(i + 1);
                }

                // Calculate median using the built‑in worksheet function.
                // CalculateFormula returns an object; convert it safely to double.
                object medianResult = sheet.CalculateFormula("MEDIAN(A1:A10)");
                double median = Convert.ToDouble(medianResult);

                // Place the result in B1.
                sheet.Cells["B1"].PutValue(median);

                // Define output file path.
                string outputPath = "MedianExample.xlsx";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
