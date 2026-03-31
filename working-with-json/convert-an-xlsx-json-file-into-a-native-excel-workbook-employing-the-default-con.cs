using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that contains Excel data
            string jsonSourcePath = "input.json";

            // Path for the resulting native Excel workbook (XLSX)
            string excelOutputPath = "output.xlsx";

            // Load options for JSON – using default settings
            JsonLoadOptions loadOptions = new JsonLoadOptions();

            // Load the JSON file into a Workbook instance
            Workbook workbook = new Workbook(jsonSourcePath, loadOptions);

            // Save the workbook as a native Excel file (XLSX) using default parameters
            workbook.Save(excelOutputPath);

            Console.WriteLine($"Conversion completed: '{jsonSourcePath}' → '{excelOutputPath}'");
        }
    }
}