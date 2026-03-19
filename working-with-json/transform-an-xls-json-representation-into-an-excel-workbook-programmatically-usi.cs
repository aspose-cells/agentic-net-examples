using System;
using Aspose.Cells;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source JSON file that represents an Excel workbook
            string jsonFilePath = "input.json";

            // Path for the output Excel file (XLSX format)
            string excelFilePath = "output.xlsx";

            // Create load options for JSON; default settings are sufficient
            JsonLoadOptions loadOptions = new JsonLoadOptions();

            // Load the JSON representation into a Workbook instance
            Workbook workbook = new Workbook(jsonFilePath, loadOptions);

            // Save the workbook as an Excel file using the default format inferred from the file extension
            workbook.Save(excelFilePath);

            Console.WriteLine($"Conversion completed: '{jsonFilePath}' -> '{excelFilePath}'");
        }
    }
}