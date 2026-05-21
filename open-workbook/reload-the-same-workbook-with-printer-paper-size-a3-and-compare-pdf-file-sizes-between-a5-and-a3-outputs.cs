using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPaperSizeComparison
{
    class Program
    {
        static void Main()
        {
            // Paths for temporary and output files
            string tempXlsxPath = "tempWorkbook.xlsx";
            string pdfA5Path = "output_A5.pdf";
            string pdfA3Path = "output_A3.pdf";

            // -------------------------------------------------
            // 1. Create a workbook, add sample data and save as XLSX
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data to make the PDF content non‑trivial
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Save the workbook to a temporary XLSX file (lifecycle: create → save)
            workbook.Save(tempXlsxPath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // 2. Load the workbook, set paper size to A5 and save as PDF
            // -------------------------------------------------
            LoadOptions loadOptionsA5 = new LoadOptions();
            loadOptionsA5.SetPaperSize(PaperSizeType.PaperA5);
            Workbook wbA5 = new Workbook(tempXlsxPath, loadOptionsA5);

            // Ensure the workbook settings reflect A5 (optional, but demonstrates the property)
            wbA5.Settings.PaperSize = PaperSizeType.PaperA5;

            // Save as PDF with A5 paper size
            wbA5.Save(pdfA5Path, SaveFormat.Pdf);

            // -------------------------------------------------
            // 3. Reload the same workbook, set paper size to A3 and save as PDF
            // -------------------------------------------------
            LoadOptions loadOptionsA3 = new LoadOptions();
            loadOptionsA3.SetPaperSize(PaperSizeType.PaperA3);
            Workbook wbA3 = new Workbook(tempXlsxPath, loadOptionsA3);

            // Apply A3 paper size to workbook settings
            wbA3.Settings.PaperSize = PaperSizeType.PaperA3;

            // Save as PDF with A3 paper size
            wbA3.Save(pdfA3Path, SaveFormat.Pdf);

            // -------------------------------------------------
            // 4. Compare file sizes
            // -------------------------------------------------
            long sizeA5 = new FileInfo(pdfA5Path).Length;
            long sizeA3 = new FileInfo(pdfA3Path).Length;

            Console.WriteLine($"PDF size with A5 paper: {sizeA5} bytes");
            Console.WriteLine($"PDF size with A3 paper: {sizeA3} bytes");

            if (sizeA3 > sizeA5)
                Console.WriteLine("A3 PDF is larger than A5 PDF.");
            else if (sizeA3 < sizeA5)
                Console.WriteLine("A3 PDF is smaller than A5 PDF.");
            else
                Console.WriteLine("Both PDFs have the same size.");

            // Cleanup temporary XLSX if desired
            // File.Delete(tempXlsxPath);
        }
    }
}