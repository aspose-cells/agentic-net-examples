using System;
using Aspose.Cells;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an XLTX workbook
            string jsonFile = "source.json";

            // Path for the resulting Excel workbook (XLSX)
            string excelFile = "result.xlsx";

            // Create load options for JSON format (default settings)
            JsonLoadOptions loadOptions = new JsonLoadOptions();

            // Load the JSON file into a Workbook instance using the load options
            Workbook workbook = new Workbook(jsonFile, loadOptions);

            // Save the workbook as a regular Excel file using default save options
            workbook.Save(excelFile, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: '{jsonFile}' -> '{excelFile}'");
        }
    }
}