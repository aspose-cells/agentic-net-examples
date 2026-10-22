// Title: C# – Load worksheets at indexes 0 and 2 with Aspose.Cells LoadOptions and export each to a separate PDF file
// Common Searches: Aspose.Cells load only selected worksheets by index and convert to PDF in C# | How to export specific Excel sheets (e.g., sheet 0 and sheet 2) to PDF using Aspose.Cells .NET | C# copy a single worksheet to a new workbook and save as PDF with Aspose.Cells | LoadOptions for selective sheet loading from an XLSX file and PDF conversion in .NET
// Tags: load specific worksheet indexes Aspose.Cells | export worksheet to PDF C# | copy worksheet to new workbook Aspose.Cells | selective sheet loading LoadOptions | single-sheet PDF conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies that the source XLSX file exists, loads it with LoadOptions (Xlsx format), iterates over the worksheet indexes 0 and 2, creates a new workbook containing only the selected sheet, and saves each workbook as a separate PDF (Sheet0.pdf, Sheet2.pdf). It includes range checks for sheet indexes and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the workbook with LoadOptions (specify the format if needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook sourceWorkbook = new Workbook(inputFile, loadOptions);

            // Indexes of worksheets to export
            int[] sheetIndexes = { 0, 2 };

            foreach (int index in sheetIndexes)
            {
                // Ensure the index is within the range of worksheets
                if (index < 0 || index >= sourceWorkbook.Worksheets.Count)
                {
                    Console.WriteLine($"Warning: Sheet index {index} is out of range. Skipping.");
                    continue;
                }

                // Create a new workbook that will contain only the selected sheet
                Workbook singleSheetWorkbook = new Workbook();
                // Remove the default empty sheet
                singleSheetWorkbook.Worksheets.Clear();

                // Copy the desired sheet from the source workbook by name
                string sheetName = sourceWorkbook.Worksheets[index].Name;
                singleSheetWorkbook.Worksheets.AddCopy(sheetName);

                // Export the single-sheet workbook to a PDF file
                string pdfFileName = $"Sheet{index}.pdf";
                singleSheetWorkbook.Save(pdfFileName, SaveFormat.Pdf);
                Console.WriteLine($"Saved '{pdfFileName}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
