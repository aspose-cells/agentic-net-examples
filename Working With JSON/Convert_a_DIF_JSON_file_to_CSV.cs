using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DifJsonToCsvConverter
    {
        public static void Run()
        {
            // Paths for input JSON (exported from DIF) and output CSV
            string jsonFilePath = "input.json";
            string csvOutputPath = "output.csv";

            // Read the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Load JSON into a workbook
            var loadOptions = new LoadOptions(LoadFormat.Json);
            using (var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonContent)))
            {
                Workbook workbook = new Workbook(jsonStream, loadOptions);

                // Save the populated workbook as CSV
                workbook.Save(csvOutputPath, SaveFormat.Csv);
            }

            Console.WriteLine($"Conversion completed. CSV saved to: {csvOutputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DifJsonToCsvConverter.Run();
        }
    }
}