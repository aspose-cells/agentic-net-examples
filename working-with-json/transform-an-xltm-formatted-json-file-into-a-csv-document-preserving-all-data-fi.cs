using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace XLTMJsonToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonFilePath = "input.json";
            string csvOutputPath = "output.csv";

            string jsonContent = File.ReadAllText(jsonFilePath);

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

            workbook.Save(csvOutputPath, SaveFormat.Csv);

            Console.WriteLine($"Conversion completed. CSV saved to: {csvOutputPath}");
        }
    }
}