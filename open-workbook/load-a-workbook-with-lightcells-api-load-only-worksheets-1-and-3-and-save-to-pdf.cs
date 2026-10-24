// Title: Load an Excel workbook, keep only the first and third worksheets, and export to PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load workbook and save only selected sheets as PDF | How to remove all worksheets except specific ones before PDF conversion using Aspose.Cells | Aspose.Cells filter worksheets by index and export to PDF in .NET | Save first and third worksheet to PDF with Aspose.Cells C# | Export selected sheets from XLSX to PDF using Aspose.Cells library
// Tags: load workbook Aspose.Cells C# | remove worksheets by index Aspose.Cells | export selected sheets to PDF Aspose.Cells | worksheet filtering Aspose.Cells .NET | PDF conversion specific worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX file with Aspose.Cells, removes every worksheet except the first and third (indexes 0 and 2), and saves the resulting workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Path for the resulting PDF
            string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (full load)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Keep only the first and third worksheets (indexes 0 and 2)
            // Remove other worksheets in reverse order to avoid index shifting
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (i != 0 && i != 2)
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the resulting workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
