using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    public class ValidationExample
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 1. Whole number validation (between 10 and 100) on cell A1
            CellArea wholeNumberArea = CellArea.CreateCellArea(0, 0, 0, 0); // A1
            int wholeNumberIndex = sheet.Validations.Add(wholeNumberArea);
            Validation wholeNumberValidation = sheet.Validations[wholeNumberIndex];
            wholeNumberValidation.Type = ValidationType.WholeNumber;
            wholeNumberValidation.Operator = OperatorType.Between;
            wholeNumberValidation.Formula1 = "10";
            wholeNumberValidation.Formula2 = "100";
            wholeNumberValidation.InputMessage = "Enter a whole number between 10 and 100.";
            wholeNumberValidation.ErrorMessage = "The value must be a whole number between 10 and 100.";
            wholeNumberValidation.ShowInput = true;
            wholeNumberValidation.ShowError = true;

            // 2. List validation on cell B1 (optional example)
            CellArea listArea = CellArea.CreateCellArea(0, 1, 0, 1); // B1
            int listIndex = sheet.Validations.Add(listArea);
            Validation listValidation = sheet.Validations[listIndex];
            listValidation.Type = ValidationType.List;
            listValidation.InCellDropDown = true;
            listValidation.Formula1 = "\"Option1,Option2,Option3\"";
            listValidation.InputMessage = "Select an option from the list.";
            listValidation.ErrorMessage = "Invalid selection.";
            listValidation.ShowInput = true;
            listValidation.ShowError = true;

            // Save the workbook
            workbook.Save("ValidationDemo.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ValidationExample.Run();
        }
    }
}