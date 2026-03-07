using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;

namespace AsposeCellsExamples
{
    public class DifCsvToJsonConverter
    {
        public static void Run()
        {
            // Path to the source DIF (or CSV with .dif extension) file
            string difFilePath = "source.dif";

            // Ensure the source file exists; create a simple sample if it does not
            if (!File.Exists(difFilePath))
            {
                string[] sampleLines =
                {
                    "Name,Age,Country",
                    "Alice,30,USA",
                    "Bob,25,Canada",
                    "Charlie,35,UK"
                };
                File.WriteAllLines(difFilePath, sampleLines);
            }

            // Load the DIF/CSV file using DifLoadOptions
            DifLoadOptions loadOptions = new DifLoadOptions();
            Workbook workbook = new Workbook(difFilePath, loadOptions);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            // Save the workbook (first worksheet) to JSON
            string jsonFilePath = "output.json";
            workbook.Save(jsonFilePath, jsonOptions);

            Console.WriteLine($"DIF file '{difFilePath}' has been converted to JSON and saved as '{jsonFilePath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DifCsvToJsonConverter.Run();
        }
    }
}