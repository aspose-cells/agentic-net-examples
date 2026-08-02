using System;
using Aspose.Cells;

namespace AsposeCellsConditionalWarningDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample numeric data in column A (cells A1:A5)
            worksheet.Cells["A1"].PutValue(5);
            worksheet.Cells["A2"].PutValue(12);
            worksheet.Cells["A3"].PutValue(8);
            worksheet.Cells["A4"].PutValue(15);
            worksheet.Cells["A5"].PutValue(3);

            // Define the range to which the validation will be applied (A1:A5)
            CellArea validationArea = CellArea.CreateCellArea("A1", "A5");

            // Add a validation rule to the worksheet
            Validation validation = worksheet.Validations[worksheet.Validations.Add()];
            validation.AddArea(validationArea);

            // Set validation type to WholeNumber and operator to LessThan
            // This will trigger when the cell value is less than the defined minimum (e.g., 10)
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.LessThan;
            validation.Formula1 = "10"; // Minimum threshold

            // Configure the warning style (not a stop error, just a warning)
            validation.AlertStyle = ValidationAlertType.Warning;

            // Set the message that will be displayed when the condition is met
            validation.ErrorTitle = "Value Too Low";
            validation.ErrorMessage = "The entered number is below the minimum allowed value of 10.";
            validation.ShowError = true; // Ensure the warning is shown

            // Optionally, show an input message when the cell is selected
            validation.ShowInput = true;
            validation.InputTitle = "Enter Value";
            validation.InputMessage = "Please enter a number greater than or equal to 10.";

            // Save the workbook to a file
            workbook.Save("ConditionalWarningDemo.xlsx");
        }
    }
}