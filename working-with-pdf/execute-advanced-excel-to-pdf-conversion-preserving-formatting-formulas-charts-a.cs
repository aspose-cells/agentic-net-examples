using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExcelToPdfConverter
    {
        public static void Convert(string excelPath, string pdfPath)
        {
            var workbook = new Workbook(excelPath);
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }

        public static void Run()
        {
            string sourceFile = "input.xlsx";
            string outputFile = "output.pdf";

            Convert(sourceFile, outputFile);
            Console.WriteLine("Excel file successfully converted to PDF.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ExcelToPdfConverter.Run();
        }
    }
}