using System;
using System.IO;
using Aspose.Cells;

namespace EmailValidationDemoApp
{
    public class EmailValidationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the validation range (cell A1)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 0,
                    EndColumn = 0
                };

                // Add a validation object for the defined area
                int validationIndex = worksheet.Validations.Add(area);
                Validation validation = worksheet.Validations[validationIndex];

                // Set validation type to Custom and provide a formula that checks for "@" and "." in the cell value
                validation.Type = ValidationType.Custom;
                validation.Formula1 = "=AND(ISNUMBER(FIND(\"@\",A1)), ISNUMBER(FIND(\".\",A1)))";

                // Configure the input message that appears when the cell is selected
                validation.ShowInput = true;
                validation.InputTitle = "Email Required";
                validation.InputMessage = "Please enter a valid email address";

                // Configure the error message that appears when the entered value does not satisfy the formula
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Email";
                validation.ErrorMessage = "The value entered is not a valid email address";

                // Define output file path
                string outputPath = "EmailValidationDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}