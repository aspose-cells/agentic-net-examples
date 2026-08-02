// Title: C# – Aspose.Cells: Apply Whole‑Number Data Validation with Accent 5 Theme Background
// Description: Creates a workbook, adds whole‑number validation (1‑100) to cells A1:B2, shows input and error messages, applies a solid Accent5 theme background to the validated cells, and saves the file as ValidationAccent5Demo.xlsx.
// Keywords: Aspose.Cells C# data validation theme color | Accent5 background Aspose.Cells | set validation cell style Aspose.Cells .NET | whole number validation Aspose.Cells | theme color background C# Aspose.Cells | Excel theme accent color Aspose.Cells | apply solid background Aspose.Cells
// Common Searches: Aspose.Cells set validation cell background theme color | C# Aspose.Cells Accent5 background for validated cells | how to use ThemeColor Accent5 in Aspose.Cells | apply whole number validation with theme color in .NET | Aspose.Cells example validation Accent5
// Developer Intent: Add whole‑number validation to a range and highlight the input cells using the workbook’s Accent5 theme color.
// Use Cases: Enforce numeric entry limits while visually distinguishing the input area with a brand‑consistent theme color. | Generate Excel templates where validated cells are automatically highlighted using the document’s Accent5 color. | Programmatically apply the same Accent5 background style to multiple validation ranges across a workbook.
// AI Prompts: Show how to change the background to Accent3 instead of Accent5 in the validation example. | Provide a snippet that adds a list‑type validation and uses a theme‑based background color for the range. | Explain how to adjust the tint of the Accent5 background color for validated cells in Aspose.Cells.

using Aspose.Cells;

// Creates a workbook, adds whole‑number validation (1‑100) to cells A1:B2, shows input and error messages, applies a solid Accent5 theme background to the validated cells, and saves the file as ValidationAccent5Demo.xlsx.
class ValidationAccent5Demo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the cell area (A1:B2) for the validation
        CellArea area = CellArea.CreateCellArea(0, 0, 1, 1);

        // Add a data validation rule to the defined area
        ValidationCollection validations = sheet.Validations;
        int validationIndex = validations.Add(area);
        Validation validation = validations[validationIndex];
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";
        validation.Formula2 = "100";
        validation.ShowInput = true;
        validation.InputTitle = "Enter Number";
        validation.InputMessage = "Please enter a whole number between 1 and 100.";
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "Number out of range.";
        validation.AlertStyle = ValidationAlertType.Stop;

        // Create a style that uses the theme's Accent5 color for the background
        Style accentStyle = workbook.CreateStyle();
        accentStyle.Pattern = BackgroundType.Solid;
        accentStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent5, 0); // no tint adjustment

        // Apply the style to each cell in the validation area
        for (int row = area.StartRow; row <= area.EndRow; row++)
        {
            for (int col = area.StartColumn; col <= area.EndColumn; col++)
            {
                cells[row, col].SetStyle(accentStyle);
            }
        }

        // Save the workbook
        workbook.Save("ValidationAccent5Demo.xlsx");
    }
}
