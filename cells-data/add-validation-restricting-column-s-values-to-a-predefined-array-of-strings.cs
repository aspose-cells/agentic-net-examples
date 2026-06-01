using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Predefined list of allowed strings for column S
        string[] allowedValues = new string[] { "Apple", "Banana", "Cherry" };
        string list = string.Join(",", allowedValues); // "Apple,Banana,Cherry"

        // Define the validation area: column S (index 18), rows 0 to 99
        CellArea validationArea = CellArea.CreateCellArea(0, 18, 99, 18);

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(validationArea);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation as a list with a dropdown
        validation.Type = ValidationType.List;
        validation.Formula1 = list;               // Set allowed values
        validation.InCellDropDown = true;         // Show dropdown in cells
        validation.AlertStyle = ValidationAlertType.Stop;
        validation.ErrorTitle = "Invalid entry";
        validation.ErrorMessage = "Please select a value from the predefined list.";
        validation.ShowError = true;

        // Save the workbook
        workbook.Save("ColumnSValidation.xlsx");
    }
}