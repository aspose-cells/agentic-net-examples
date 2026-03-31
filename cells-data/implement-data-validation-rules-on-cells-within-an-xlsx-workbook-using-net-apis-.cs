using System;
using Aspose.Cells;

class DataValidationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Whole number validation (between 1 and 100) for cells A1:A5
        // -------------------------------------------------
        // Define the cell area A1:A5 (rows 0-4, column 0)
        CellArea wholeNumberArea = CellArea.CreateCellArea(0, 0, 4, 0);
        // Add a validation for this area and retrieve the Validation object
        int wholeNumberIdx = sheet.Validations.Add(wholeNumberArea);
        Validation wholeNumberValidation = sheet.Validations[wholeNumberIdx];
        // Configure validation properties
        wholeNumberValidation.Type = ValidationType.WholeNumber;
        wholeNumberValidation.Operator = OperatorType.Between;
        wholeNumberValidation.Formula1 = "1";
        wholeNumberValidation.Formula2 = "100";
        wholeNumberValidation.InputMessage = "Enter a whole number between 1 and 100.";
        wholeNumberValidation.ErrorMessage = "Invalid entry. Must be a whole number 1‑100.";
        wholeNumberValidation.ShowInput = true;
        wholeNumberValidation.ShowError = true;

        // -------------------------------------------------
        // 2. List validation (drop‑down) for cell B1
        // -------------------------------------------------
        CellArea listArea = CellArea.CreateCellArea(0, 1, 0, 1); // B1
        int listIdx = sheet.Validations.Add(listArea);
        Validation listValidation = sheet.Validations[listIdx];
        listValidation.Type = ValidationType.List;
        listValidation.InCellDropDown = true;
        // Comma‑separated list of allowed values
        listValidation.Formula1 = "Red,Green,Blue";
        listValidation.InputMessage = "Select a color.";
        listValidation.ErrorMessage = "Invalid color selected.";
        listValidation.ShowInput = true;
        listValidation.ShowError = true;

        // -------------------------------------------------
        // 3. Date validation (between two dates) for cells C1:C3
        // -------------------------------------------------
        CellArea dateArea = CellArea.CreateCellArea(0, 2, 2, 2); // C1:C3
        int dateIdx = sheet.Validations.Add(dateArea);
        Validation dateValidation = sheet.Validations[dateIdx];
        dateValidation.Type = ValidationType.Date;
        dateValidation.Operator = OperatorType.Between;
        // Use Excel DATE function in A1‑style formula
        dateValidation.Formula1 = "DATE(2023,1,1)";
        dateValidation.Formula2 = "DATE(2023,12,31)";
        dateValidation.InputMessage = "Enter a date in the year 2023.";
        dateValidation.ErrorMessage = "Date must be within 2023.";
        dateValidation.ShowInput = true;
        dateValidation.ShowError = true;

        // -------------------------------------------------
        // 4. Custom validation using a formula for cell D1
        // -------------------------------------------------
        CellArea customArea = CellArea.CreateCellArea(0, 3, 0, 3); // D1
        int customIdx = sheet.Validations.Add(customArea);
        Validation customValidation = sheet.Validations[customIdx];
        customValidation.Type = ValidationType.Custom;
        // Example: length of text in A1 must be 5 characters or less
        customValidation.Formula1 = "=LEN(A1)<=5";
        customValidation.InputMessage = "Length of A1 must be ≤5 characters.";
        customValidation.ErrorMessage = "Invalid length in A1.";
        customValidation.ShowInput = true;
        customValidation.ShowError = true;

        // -------------------------------------------------
        // Save the workbook with all validations applied
        // -------------------------------------------------
        workbook.Save("DataValidationDemo.xlsx");
    }
}