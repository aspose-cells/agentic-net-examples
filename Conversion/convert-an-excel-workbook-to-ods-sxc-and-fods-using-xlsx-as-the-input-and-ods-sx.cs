using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertXlsxToMultipleFormats.Run();
        }
    }

    public class ConvertXlsxToMultipleFormats
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            Workbook workbook = new Workbook(sourcePath);
            workbook.Save("output.ods", SaveFormat.Ods);
            workbook.Save("output.sxc", SaveFormat.Sxc);
            workbook.Save("output.fods", SaveFormat.Fods);
            Console.WriteLine("Conversion completed: ODS, SXC, and FODS files have been created.");
        }
    }
}