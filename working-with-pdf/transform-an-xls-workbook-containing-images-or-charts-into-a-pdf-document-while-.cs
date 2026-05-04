using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            XlsToPdfConverter.Run();
        }
    }

    public class XlsToPdfConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xls";
            string destPath = "output.pdf";

            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Excel file successfully converted to PDF.");
        }
    }
}