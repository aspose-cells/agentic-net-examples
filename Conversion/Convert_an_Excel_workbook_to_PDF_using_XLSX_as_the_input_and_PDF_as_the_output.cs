using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertExcelToPdf
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired path for the output PDF file
            string destPath = "output.pdf";

            // Load the Excel workbook
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as PDF
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Excel file has been successfully converted to PDF.");
        }
    }
}