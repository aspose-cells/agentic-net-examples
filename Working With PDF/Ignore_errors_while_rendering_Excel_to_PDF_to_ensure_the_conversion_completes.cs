using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsIgnoreErrorDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            using (MemoryStream ms = new MemoryStream())
            {
                sheet.Shapes.AddPicture(0, 0, 5, 5, ms);
            }

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                IgnoreError = true
            };

            string outputPath = "Output_IgnoringErrors.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF with IgnoreError enabled: {outputPath}");
        }
    }
}