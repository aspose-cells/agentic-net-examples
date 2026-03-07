using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            CsvToJsonConverter.Run();
        }
    }

    public class CsvToJsonConverter
    {
        public static void Run()
        {
            string csvPath = "input.csv";
            string jsonPath = "output.json";

            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"CSV file \"{csvPath}\" has been successfully converted to JSON file \"{jsonPath}\".");
        }
    }
}