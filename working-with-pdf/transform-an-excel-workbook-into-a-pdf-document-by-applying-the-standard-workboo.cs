using System;
using Aspose.Cells; // Namespace for Workbook and SaveFormat

namespace AsposeCellsPdfConversion
{
    public class WorkbookToPdf
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Desired PDF output path
            string pdfPath = "output.pdf";

            // Load the Excel workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Save the loaded workbook as a PDF document
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
        }
    }
}