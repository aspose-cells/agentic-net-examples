using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsxToPdfConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.pdf";

            Workbook workbook = new Workbook(sourcePath);
            workbook.CalculateFormula();
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Workbook successfully converted to PDF.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XlsxToPdfConverter.Run();
        }
    }
}