using System;
using Aspose.Cells;

namespace AsposeCellsValidationWithThemeBackground
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the cell (e.g., A1) that will have the validation
            string targetCellName = "A1";

            // -------------------------------------------------
            // 1. Create a style that uses the theme's Accent5 color
            // -------------------------------------------------
            Style themeStyle = workbook.CreateStyle();
            // Set a solid fill pattern so the background color is visible
            themeStyle.Pattern = BackgroundType.Solid;
            // Use Accent5 with no tint (0) – you can adjust the tint if needed
            themeStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent5, 0);

            // Apply the style to the target cell
            cells[targetCellName].SetStyle(themeStyle);
            cells[targetCellName].PutValue("Enter a number between 10 and 100");

            // -------------------------------------------------
            // 2. Add a data validation rule to the same cell
            // -------------------------------------------------
            // Define the validation area (single cell A1)
            CellArea validationArea = CellArea.CreateCellArea(0, 0, 0, 0); // Row 0, Column 0

            // Add the validation to the worksheet
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];

            // Configure the validation (whole number between 10 and 100)
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "100";

            // Show the input message when the cell is selected
            validation.ShowInput = true;
            validation.InputTitle = "Number Required";
            validation.InputMessage = "Please enter a whole number between 10 and 100.";

            // Show an error message if the input is invalid
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "The value must be a whole number between 10 and 100.";
            validation.AlertStyle = ValidationAlertType.Stop;

            // -------------------------------------------------
            // 3. Save the workbook
            // -------------------------------------------------
            workbook.Save("ValidationWithAccent5Background.xlsx");
        }
    }
}