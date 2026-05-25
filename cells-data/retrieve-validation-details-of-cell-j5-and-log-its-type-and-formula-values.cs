using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the cell area for J5 (start and end are the same)
                CellArea area = CellArea.CreateCellArea("J5", "J5");

                // Add validation to the defined cell area
                int validationIndex = worksheet.Validations.Add(area);
                Validation demoValidation = worksheet.Validations[validationIndex];
                demoValidation.Type = ValidationType.WholeNumber;
                demoValidation.Operator = OperatorType.Between;
                demoValidation.Formula1 = "10";
                demoValidation.Formula2 = "20";

                // Retrieve the validation applied to J5
                Cell cell = worksheet.Cells["J5"];
                Validation validation = cell.GetValidation();

                if (validation != null)
                {
                    Console.WriteLine("Validation Type: " + validation.Type);
                    Console.WriteLine("Formula1: " + validation.Formula1);
                    Console.WriteLine("Formula2: " + validation.Formula2);
                }
                else
                {
                    Console.WriteLine("No validation applied to cell J5.");
                }

                // Save the workbook (optional)
                string outputPath = "ValidationJ5.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}