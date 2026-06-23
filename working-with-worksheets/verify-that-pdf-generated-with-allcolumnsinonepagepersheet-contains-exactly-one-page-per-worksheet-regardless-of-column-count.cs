using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate each sheet with many columns to exceed normal page width
            PopulateSheetWithColumns(sheet1, 120);
            PopulateSheetWithColumns(sheet2, 150);

            // Configure PDF save options to force all columns onto a single page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true
            };

            // Save the workbook as PDF
            string pdfPath = "AllColumnsOnePagePerSheet.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF saved to {pdfPath}");

            // Verify that each worksheet renders to exactly one page
            VerifyOnePagePerWorksheet(workbook);
        }

        // Adds sample data across the specified number of columns
        private static void PopulateSheetWithColumns(Worksheet sheet, int columnCount)
        {
            for (int col = 0; col < columnCount; col++)
            {
                // Header
                sheet.Cells[0, col].PutValue($"Col {col + 1}");
                // Sample data rows
                for (int row = 1; row <= 10; row++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                }
            }
        }

        // Uses SheetPrintingPreview to evaluate page count for each worksheet
        private static void VerifyOnePagePerWorksheet(Workbook workbook)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Rendering options must match the save options used for PDF generation
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    AllColumnsInOnePagePerSheet = true,
                    OnePagePerSheet = true
                };

                SheetPrintingPreview preview = new SheetPrintingPreview(sheet, renderOptions);
                int pageCount = preview.EvaluatedPageCount;

                Console.WriteLine($"Worksheet \"{sheet.Name}\" page count: {pageCount}");
                if (pageCount != 1)
                {
                    Console.WriteLine("Verification failed: Expected exactly one page per worksheet.");
                }
                else
                {
                    Console.WriteLine("Verification succeeded.");
                }
            }
        }
    }
}