using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class ExcelToJsonConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destinationPath = "output.json";

            Workbook workbook = new Workbook(sourcePath);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,
                ExportEmptyCells = true
            };

            workbook.Save(destinationPath, jsonOptions);

            Console.WriteLine($"Excel workbook '{sourcePath}' has been successfully converted to JSON at '{destinationPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToJsonConverter.Run();
        }
    }
}