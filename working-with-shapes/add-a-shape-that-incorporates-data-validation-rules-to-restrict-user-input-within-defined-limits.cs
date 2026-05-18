using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data that will be used as the list box source
            for (int i = 0; i < 5; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Add a ListBox shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape listBox = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Set the range that fills the list box (A1:A5)
            listBox.SetInputRange("$A$1:$A$5", false, false);

            // Link the selected value of the list box to cell B2
            listBox.SetLinkedCell("$B$2", false, false);

            // Define a validation area for the linked cell (B2)
            CellArea validationArea = CellArea.CreateCellArea(1, 1, 1, 1); // Row 2, Column 2 (B2)

            // Add a validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation to allow whole numbers between 10 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "100";

            // Set user-friendly messages
            validation.InputMessage = "Enter a number between 10 and 100.";
            validation.InputTitle = "Number Input";
            validation.ErrorMessage = "The value must be a whole number between 10 and 100.";
            validation.ErrorTitle = "Invalid Input";

            // Show messages when the cell is selected or when invalid data is entered
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            workbook.Save("ShapeWithValidation.xlsx");
        }
    }
}