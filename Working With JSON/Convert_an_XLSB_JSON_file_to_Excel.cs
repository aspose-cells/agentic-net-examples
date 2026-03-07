using System;
using Aspose.Cells;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the JSON file exported from an XLSB workbook
            string jsonFilePath = "input.json";

            // Desired output Excel file (XLSX format)
            string outputExcelPath = "output.xlsx";

            // Load the JSON file into a workbook using Aspose.Cells JSON support
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);
            Workbook workbook = new Workbook(jsonFilePath, loadOptions);

            // Save the workbook as XLSX
            workbook.Save(outputExcelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Excel file saved to '{outputExcelPath}'.");
        }
    }
}