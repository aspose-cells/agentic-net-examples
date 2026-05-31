using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBatchAutoFitRows
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Auto‑fit all rows in the current worksheet
                sheet.AutoFitRows();

                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,                 // each sheet fits on one page
                    AllColumnsInOnePagePerSheet = true      // all columns on one page
                };

                // Render only the current worksheet
                pdfOptions.SheetSet = new SheetSet(new int[] { i });

                // Define output PDF file name for the current sheet
                string outputPath = $"Sheet{i + 1}.pdf";

                // Save the workbook (only the selected sheet) as PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Saved worksheet '{sheet.Name}' as PDF: {outputPath}");
            }
        }
    }
}