using System;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class ConvertExcelToPdfDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.pdf";

            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}