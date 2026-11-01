// Title: C# code to delete all pivot tables from an Excel workbook and export it as a PDF with Aspose.Cells
// Common Searches: C# how to delete every pivot table in an Excel file before converting to PDF using Aspose.Cells | Aspose.Cells remove pivot tables from workbook then save as PDF in .NET | programmatically clear pivot tables from worksheets and export to PDF with Aspose.Cells C# | remove all pivot tables from Excel workbook using Aspose.Cells API and generate PDF report | Aspose.Cells .NET example for cleaning workbook of pivot tables prior to PDF conversion
// Tags: Aspose.Cells remove pivot tables C# | Aspose.Cells export to PDF .NET | clean Excel workbook for PDF generation | pivot table deletion before PDF conversion | Workbook.Save as PDF after cleanup

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, iterates each worksheet, removes all pivot tables by looping backwards through the PivotTables collection, and then saves the cleaned workbook as a PDF using Aspose.Cells, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "report.pdf";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all pivot tables from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate backwards to safely remove items from the collection
                for (int i = sheet.PivotTables.Count - 1; i >= 0; i--)
                {
                    sheet.PivotTables.RemoveAt(i);
                }
            }

            // Save the cleaned workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
