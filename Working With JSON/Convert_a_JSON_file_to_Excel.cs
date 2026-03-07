using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class JsonToExcelConverter
    {
        public static void Run()
        {
            // Paths for source JSON file and destination Excel file
            string jsonPath = "input.json";
            string excelPath = "output.xlsx";

            try
            {
                // Create load options for JSON (default options are sufficient for most cases)
                JsonLoadOptions loadOptions = new JsonLoadOptions();

                // Load the JSON file into a Workbook instance using the load options
                Workbook workbook = new Workbook(jsonPath, loadOptions);

                // Save the workbook as an Excel file (XLSX format)
                workbook.Save(excelPath, SaveFormat.Xlsx);

                Console.WriteLine($"Conversion completed successfully: '{jsonPath}' → '{excelPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            JsonToExcelConverter.Run();
        }
    }
}