using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation area for column G (zero‑based index 6), rows 0‑1000
        CellArea area = CellArea.CreateCellArea(0, 6, 1000, 6);

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Set validation to allow whole numbers between 10 and 500
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "10";
        validation.Formula2 = "500";

        // Optional user prompts
        validation.InputTitle = "Enter Integer";
        validation.InputMessage = "Please enter an integer between 10 and 500.";
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "Value must be an integer between 10 and 500.";
        validation.ShowInput = true;
        validation.ShowError = true;

        // Save the workbook
        workbook.Save("ColumnGValidation.xlsx");
    }
}