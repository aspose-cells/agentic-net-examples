using System;
using Aspose.Cells;

namespace JsonToExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the JSON file that was exported from an XLSB workbook
            string jsonFilePath = "input.json";

            // Desired output Excel file (native .xlsx format)
            string excelOutputPath = "output.xlsx";

            // Load the JSON file as a workbook.
            // LoadOptions with LoadFormat.Json tells Aspose.Cells to interpret the source as JSON.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);

            // Initialize the Workbook using the constructor that accepts a file path and LoadOptions.
            Workbook workbook = new Workbook(jsonFilePath, loadOptions);

            // Save the workbook in native Excel format.
            // Using the Save method that accepts a file name and SaveFormat.
            workbook.Save(excelOutputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: '{jsonFilePath}' -> '{excelOutputPath}'");
        }
    }
}