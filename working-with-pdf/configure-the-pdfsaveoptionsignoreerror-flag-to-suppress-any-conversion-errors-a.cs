using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            PdfIgnoreErrorDemo.Run();
        }
    }

    public class PdfIgnoreErrorDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("input.xlsx");
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.IgnoreError = true;
            workbook.Save("output.pdf", pdfOptions);
            Console.WriteLine("PDF saved with IgnoreError enabled.");
        }
    }
}