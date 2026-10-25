// Title: Refresh all pivot tables in an Excel workbook and export the result to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# code to refresh pivot tables in an Excel file before saving as PDF with Aspose.Cells | How to recalculate pivot table data and export workbook to PDF using Aspose.Cells .NET | Aspose.Cells example for updating pivot tables and generating PDF output | Export refreshed pivot tables to PDF while preserving layout in C# | Saving an Excel workbook with refreshed pivot tables to PDF using Aspose.Cells API
// Tags: Aspose.Cells update pivot cache C# | PDF conversion preserving pivot layout Aspose.Cells | calculate pivot data before PDF save | save refreshed Excel workbook as PDF .NET | Aspose.Cells pivot table layout retention PDF

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing Excel workbook, checks the file's existence, iterates through each worksheet to refresh and recalculate every pivot table, recalculates the whole workbook to update dependent formulas, and finally saves the refreshed workbook as a PDF while preserving the pivot table layout and calculated values. Exceptions are caught and logged.
class ExportPivotTableToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Refresh all pivot tables in the workbook and recalculate their values
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Refresh the data source of the pivot table
                    pivot.RefreshData();

                    // Recalculate the pivot table data
                    pivot.CalculateData();
                }
            }

            // Recalculate the entire workbook (covers any formulas that depend on the pivot data)
            workbook.CalculateFormula();

            // Save the refreshed workbook as a PDF, preserving layout and calculated values
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Successfully exported pivot tables to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
