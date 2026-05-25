using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DateValidationInColumnN
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define validation area for column N (index 13) rows 1‑1000 (0‑based rows 0‑999)
                CellArea validationArea = CellArea.CreateCellArea(0, 13, 999, 13);

                // Add validation to the defined area
                int validationIndex = sheet.Validations.Add(validationArea);
                Validation validation = sheet.Validations[validationIndex];

                // Configure date validation between 01/01/2020 and 12/31/2025
                validation.Type = ValidationType.Date;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = new DateTime(2020, 1, 1).ToOADate().ToString();
                validation.Formula2 = new DateTime(2025, 12, 31).ToOADate().ToString();

                // Optional user‑friendly messages
                validation.InputMessage = "Enter a date between 01/01/2020 and 12/31/2025.";
                validation.InputTitle = "Date Input";
                validation.ErrorMessage = "The date is out of the allowed range.";
                validation.ErrorTitle = "Invalid Date";
                validation.ShowInput = true;
                validation.ShowError = true;

                // Save the workbook
                string outputPath = "DateValidation_ColumnN.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DateValidationInColumnN.Run();
        }
    }
}