// Title: Export worksheets at indexes 0 and 2 to separate PDF files with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook using Aspose.Cells, extracts worksheets with indexes 0 and 2 into individual workbooks, and saves each as a separate PDF file. | Create a reusable C# method that takes a source Excel path and an array of worksheet indexes, then uses Aspose.Cells to export each specified sheet to its own PDF document.
// Common Searches: Aspose.Cells C# export only selected worksheet indexes to PDF | How to save individual Excel sheets as separate PDF files using Aspose.Cells .NET | C# copy a specific worksheet to a new workbook and convert to PDF with Aspose.Cells | Load workbook with limited sheets using Aspose.Cells and generate PDFs per sheet | Export sheet 0 and sheet 2 from Excel to PDF using Aspose.Cells API
// Tags: export specific worksheets to PDF Aspose.Cells | load workbook with selected sheet indexes C# | copy worksheet to new workbook Aspose.Cells | save worksheet as PDF Aspose.Cells | Aspose.Cells selective sheet export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExport
{
    // The example checks for the input.xlsx file, loads it with Aspose.Cells, selects worksheets at indexes 0 and 2, copies each into a temporary workbook, and saves each temporary workbook as a separate PDF (Sheet0.pdf, Sheet2.pdf), handling out‑of‑range indexes and runtime errors.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook file
                const string inputFile = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                    return;
                }

                // Define which worksheets to export (indexes 0 and 2)
                int[] sheetsToExport = new int[] { 0, 2 };

                // Load the workbook (all sheets are loaded)
                Workbook workbook = new Workbook(inputFile);

                // Export each requested worksheet to a separate PDF file
                foreach (int sheetIndex in sheetsToExport)
                {
                    try
                    {
                        // Ensure the sheet index is within the actual worksheet count
                        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
                        {
                            Console.WriteLine($"Warning: Sheet index {sheetIndex} is out of range. Skipping.");
                            continue;
                        }

                        // Create a temporary workbook containing only the desired sheet
                        Workbook sheetWorkbook = new Workbook();
                        sheetWorkbook.Worksheets.Clear();

                        // Copy the specific worksheet by name into the new workbook
                        string sourceSheetName = workbook.Worksheets[sheetIndex].Name;
                        sheetWorkbook.Worksheets.AddCopy(sourceSheetName);

                        // Build output PDF file name
                        string pdfFileName = $"Sheet{sheetIndex}.pdf";

                        // Save the single sheet as a PDF
                        sheetWorkbook.Save(pdfFileName, SaveFormat.Pdf);
                        Console.WriteLine($"Saved sheet {sheetIndex} to '{pdfFileName}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error exporting sheet {sheetIndex}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
