using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            TxtToJsonConverter.Run();
        }
    }

    public class TxtToJsonConverter
    {
        public static void Run()
        {
            string txtFilePath = "input.txt";
            string jsonFilePath = "output.json";

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = '\t',
                ConvertNumericData = true,
                ConvertDateTimeData = true
            };

            cells.ImportCSV(txtFilePath, loadOptions, 0, 0);

            int totalRows = cells.MaxDataRow + 1;
            int totalCols = cells.MaxDataColumn + 1;

            Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, totalRows, totalCols);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            string jsonContent = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            File.WriteAllText(jsonFilePath, jsonContent);

            Console.WriteLine($"Conversion completed. JSON saved to: {jsonFilePath}");
        }
    }
}