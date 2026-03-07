using System;
using Aspose.Cells;

namespace AsposeCellsConversionExample
{
    class Program
    {
        static void Main(string[] args)
        {
            XlsmToJsonConverter.Run();
        }
    }

    public class XlsmToJsonConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsm";
            string outputPath = "output.json";

            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(outputPath, SaveFormat.Json);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }
}