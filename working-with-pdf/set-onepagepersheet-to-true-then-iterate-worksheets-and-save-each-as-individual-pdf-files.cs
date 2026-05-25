using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsOnePagePerSheetExample
{
    class Program
    {
        static void Main()
        {
            // Create a sample workbook with multiple worksheets
            Workbook workbook = new Workbook();
            // Add some data to the default sheet
            Worksheet sheet0 = workbook.Worksheets[0];
            sheet0.Name = "Sheet1";
            for (int i = 0; i < 5; i++)
                sheet0.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");

            // Add additional worksheets and fill them with sample data
            for (int s = 2; s <= 4; s++)
            {
                Worksheet ws = workbook.Worksheets.Add($"Sheet{s}");
                for (int i = 0; i < 5; i++)
                    ws.Cells[i, 0].PutValue($"Sheet{s} Row {i + 1}");
            }

            // Create PDF save options and enable OnePagePerSheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true; // All content of a sheet will be rendered on a single PDF page

            // Iterate through each worksheet and save it as an individual PDF file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Configure the SheetSet to include only the current worksheet (zero‑based index)
                pdfOptions.SheetSet = new SheetSet(new int[] { i });

                // Build a file name for the output PDF
                string outputFile = $"Worksheet_{i + 1}_{workbook.Worksheets[i].Name}.pdf";

                // Save the workbook using the configured options; only the selected sheet will be exported
                workbook.Save(outputFile, pdfOptions);
            }

            Console.WriteLine("All worksheets have been saved as individual PDF files with OnePagePerSheet enabled.");
        }
    }
}