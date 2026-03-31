using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class ExcelToPdfConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string outputPath = "output.pdf";

            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine("Excel file has been successfully converted to PDF.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToPdfConverter.Run();
        }
    }
}