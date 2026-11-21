// Title: Load only the first and third worksheets with Aspose.Cells LightCells API in C# and export them to PDF
// Common Searches: C# Aspose.Cells load only selected sheets and convert to PDF | How to export the first and third worksheet of an Excel file to PDF using Aspose.Cells | Remove unwanted worksheets before saving workbook as PDF with Aspose.Cells .NET | Aspose.Cells LightCells API example for selective sheet conversion to PDF in C# | Save specific Excel sheets to PDF without loading entire workbook using Aspose.Cells
// Tags: Aspose.Cells LightCells selective sheet export | C# export specific worksheets to PDF | Aspose.Cells remove worksheets before PDF conversion | LightCells API load partial workbook C# | Excel to PDF conversion with selected sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks that the input Excel file exists, loads the workbook, removes all worksheets except the first and third (indices 0 and 2), and then saves the remaining sheets directly to a PDF file using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input workbook and output PDF
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook without any custom filter
                Workbook workbook = new Workbook(inputPath);

                // Keep only worksheets with indices 0 and 2 (first and third sheets)
                // Remove other worksheets in reverse order to preserve indices
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    if (i != 0 && i != 2)
                    {
                        workbook.Worksheets.RemoveAt(i);
                    }
                }

                // Save the selected sheets to PDF
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
