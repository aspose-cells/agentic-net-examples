using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLightCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string sourceFile = "input.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the workbook with default load options
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(sourceFile, loadOptions);

                // Determine which sheet indexes are available (0‑based)
                List<int> sheetIndexes = new List<int>();
                if (workbook.Worksheets.Count > 0) sheetIndexes.Add(0);   // first sheet
                if (workbook.Worksheets.Count > 2) sheetIndexes.Add(2);   // third sheet

                // If the requested sheets are not present, fall back to all sheets
                if (sheetIndexes.Count == 0)
                {
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                        sheetIndexes.Add(i);
                }

                // Prepare PDF save options and specify the selected sheets
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    SheetSet = new SheetSet(sheetIndexes.ToArray())
                };

                // Save the selected sheets to a PDF file
                string outputPdf = "output.pdf";
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"Workbook saved to PDF with sheets {string.Join(", ", sheetIndexes)}: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}