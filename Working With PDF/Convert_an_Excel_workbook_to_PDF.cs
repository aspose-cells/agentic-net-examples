using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertExcelToPdfDemo
    {
        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path where the PDF will be saved
            string pdfPath = "output.pdf";

            // Load the Excel workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Excel file '{sourcePath}' has been successfully converted to PDF at '{pdfPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertExcelToPdfDemo.Run();
        }
    }
}