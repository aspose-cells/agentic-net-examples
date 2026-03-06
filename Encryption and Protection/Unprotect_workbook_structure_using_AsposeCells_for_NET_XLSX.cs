using System;
using Aspose.Cells;

namespace AsposeCellsUnprotectWorkbook
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the encrypted workbook
            string inputPath = "protected_workbook.xlsx";

            // Password used to open the encrypted workbook
            string password = "password123";

            // Load the workbook with the password (if required)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook without any further protection handling
            string outputPath = "unprotected_workbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook has been saved to '{outputPath}'.");
        }
    }
}