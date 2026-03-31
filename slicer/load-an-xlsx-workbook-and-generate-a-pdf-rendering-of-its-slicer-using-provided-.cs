using System;
using Aspose.Cells;

namespace SlicerPdfRenderDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file that contains slicers
            string sourcePath = "input.xlsx";

            // Path for the generated PDF file
            string pdfPath = "slicer_rendered.pdf";

            // Load the workbook from the specified file (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as PDF. The slicers present in the worksheet will be rendered in the PDF.
            // This uses the Workbook.Save(string, SaveFormat) method.
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook with slicers has been rendered to PDF at: {pdfPath}");
        }
    }
}