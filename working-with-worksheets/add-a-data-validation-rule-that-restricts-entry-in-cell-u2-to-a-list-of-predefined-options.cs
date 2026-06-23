using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDataValidationDemo
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

                // Add a validation rule for cell U2
                int validationIndex = worksheet.Validations.Add();
                Validation validation = worksheet.Validations[validationIndex];
                validation.Type = ValidationType.List;                     // Allow only listed values
                validation.Formula1 = "OptionA,OptionB,OptionC";           // Comma‑separated list
                validation.InCellDropDown = true;                         // Show drop‑down arrow

                // Define output file path
                string outputPath = "DataValidation_U2.xlsx";

                // Ensure the directory exists (prevents FileNotFoundException)
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook with the validation applied
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}