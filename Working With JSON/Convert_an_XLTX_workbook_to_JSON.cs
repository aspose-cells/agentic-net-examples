using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XltxToJsonConverter
    {
        public static void Run(string templatePath, string jsonPath)
        {
            Workbook workbook = new Workbook(templatePath);
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportNestedStructure = false,
                HasHeaderRow = true,
                ExportAsString = true,
                Indent = "  "
            };
            workbook.Save(jsonPath, jsonOptions);
            Console.WriteLine($"XLTX workbook '{templatePath}' has been successfully converted to JSON at '{jsonPath}'.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <templatePath> <jsonPath>");
                return;
            }

            string templatePath = args[0];
            string jsonPath = args[1];
            XltxToJsonConverter.Run(templatePath, jsonPath);
        }
    }
}