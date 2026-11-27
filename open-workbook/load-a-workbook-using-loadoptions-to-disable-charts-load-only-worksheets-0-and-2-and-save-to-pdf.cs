// Title: Load an Excel workbook, keep only worksheets 0 and 2, and export to PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export only selected sheets to PDF | how to remove unwanted worksheets before saving Excel as PDF using Aspose.Cells | Aspose.Cells keep sheet index 0 and 2 when converting to PDF | filter worksheets in a workbook prior to PDF export with Aspose.Cells .NET | Aspose.Cells convert specific worksheets to PDF without loading all sheets
// Tags: select specific worksheets Aspose.Cells C# | worksheet filtering before PDF conversion Aspose.Cells | Aspose.Cells save workbook as PDF C# | remove worksheets programmatically Aspose.Cells | Excel to PDF conversion with sheet selection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the input Excel file exists, loads it with Aspose.Cells, removes every worksheet except those at indexes 0 and 2, and then saves the resulting workbook as a PDF file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (full load)
            Workbook workbook = new Workbook(inputPath);

            // Keep only worksheets with indexes 0 and 2 (if they exist)
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (i != 0 && i != 2)
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
