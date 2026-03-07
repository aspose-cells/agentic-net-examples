using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsxToTiffConverter
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Save the workbook as a multi‑page TIFF file
            workbook.Save("output.tiff", SaveFormat.Tiff);

            Console.WriteLine("Workbook successfully converted to TIFF: output.tiff");
        }
    }
}