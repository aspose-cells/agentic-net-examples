using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            XlsxToOdsConverter.Run();
        }
    }

    public class XlsxToOdsConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string outputPath = "output.ods";

            ConversionUtility.Convert(sourcePath, outputPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }
}