// Title: Load an Excel workbook with Aspose.Cells, retain only named ranges containing "Total", and save as PDF using C#
// Common Searches: C# Aspose.Cells filter named ranges by keyword before PDF export | How to delete defined names that don't contain 'Total' using Aspose.Cells | Save an Excel file as PDF after removing specific named ranges with Aspose.Cells | Aspose.Cells remove unwanted named ranges case‑insensitive C# | Export filtered workbook to PDF with Aspose.Cells .NET
// Tags: named range filtering Aspose.Cells C# | defined name removal Aspose.Cells | pdf conversion Aspose.Cells .NET | case‑insensitive name check Aspose.Cells | workbook manipulation before PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (input.xlsx) using Aspose.Cells, iterates through its named ranges, removes any defined name whose text does not contain the word "Total" (case‑insensitive), ensures the output directory exists, and then saves the resulting workbook as a PDF (output.pdf) while handling load, save, and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Filter defined names (named ranges): keep only those containing "Total" (case‑insensitive)
            // In Aspose.Cells, named ranges are accessed via workbook.Worksheets.Names collection.
            for (int i = workbook.Worksheets.Names.Count - 1; i >= 0; i--)
            {
                Aspose.Cells.Name namedRange = workbook.Worksheets.Names[i];
                if (namedRange == null || string.IsNullOrEmpty(namedRange.Text))
                    continue;

                if (!namedRange.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    workbook.Worksheets.Names.RemoveAt(i);
                }
            }

            // Output PDF path
            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // Save the modified workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved as PDF to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
