using System;
using Aspose.Cells;

namespace AsposeCellsValidationInfo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or specify the required one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get cell J5
            Cell cell = worksheet.Cells["J5"];

            // Retrieve the validation applied to the cell
            Validation validation = cell.GetValidation();

            if (validation != null)
            {
                // Log validation type
                Console.WriteLine("Validation Type: " + validation.Type);

                // Log formula values (Formula1 and Formula2)
                Console.WriteLine("Formula1: " + validation.Formula1);
                Console.WriteLine("Formula2: " + validation.Formula2);
            }
            else
            {
                Console.WriteLine("No validation applied to cell J5.");
            }

            // Optionally save the workbook (if any changes were made)
            workbook.Save("output.xlsx");
        }
    }
}