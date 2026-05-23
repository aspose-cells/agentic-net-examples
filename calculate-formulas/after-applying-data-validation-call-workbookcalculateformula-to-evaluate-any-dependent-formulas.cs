using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDataValidationAndCalculate
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // Set up sample data and formulas
                // -------------------------------------------------
                // Cell A1 will hold a base value
                cells["A1"].PutValue(10);

                // Cell B1 contains a formula that depends on A1
                cells["B1"].Formula = "=A1*2";

                // -------------------------------------------------
                // Add data validation to cell A1 (must be between 5 and 20)
                // -------------------------------------------------
                // Define the area (cell A1) to which the validation applies
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 0,
                    EndColumn = 0
                };

                // Add a new validation for the defined area.
                // In older Aspose.Cells versions Add returns the index of the validation.
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];

                // Configure the validation rule
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "5";   // Minimum value
                validation.Formula2 = "20";  // Maximum value

                // -------------------------------------------------
                // After setting validation, calculate all formulas
                // -------------------------------------------------
                workbook.CalculateFormula();

                // -------------------------------------------------
                // Output the results to verify calculation
                // -------------------------------------------------
                Console.WriteLine("A1 (validated value) : " + cells["A1"].IntValue);
                Console.WriteLine("B1 (formula result) : " + cells["B1"].IntValue);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "DataValidationAndCalculate.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}