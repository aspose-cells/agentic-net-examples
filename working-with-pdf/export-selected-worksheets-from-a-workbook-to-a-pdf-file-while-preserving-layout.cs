using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExportSelectedSheetsToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Exporter.Run();
        }
    }

    public static class Exporter
    {
        public static void Run()
        {
            // Path to the source Excel workbook
            string sourceFile = "input.xlsx";

            // Path for the resulting PDF file
            string pdfFile = "selected_sheets.pdf";

            // Define which worksheets to export (zero‑based indexes)
            // Example: export the first and third worksheets
            int[] selectedSheetIndexes = { 0, 2 };

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourceFile);

            // Build a list of valid sheet names based on the requested indexes
            List<string> sheetNames = new List<string>();
            foreach (int idx in selectedSheetIndexes)
            {
                if (idx >= 0 && idx < workbook.Worksheets.Count)
                {
                    sheetNames.Add(workbook.Worksheets[idx].Name);
                }
            }

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Use SheetSet to specify the exact sheets to render
            pdfOptions.SheetSet = new SheetSet(sheetNames.ToArray());

            // Save the selected worksheets to PDF
            workbook.Save(pdfFile, pdfOptions);

            Console.WriteLine($"Selected worksheets exported to PDF: {pdfFile}");
        }
    }
}