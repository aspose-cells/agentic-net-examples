// Title: C# – Add List Validation with Accent5 Theme Background Using Aspose.Cells
// Description: Creates a workbook, defines range A1:B2, adds a list‑type validation (Yes/No), applies a solid style whose background uses the workbook’s Accent5 theme color, and saves the file as ValidationWithAccent5Background.xlsx.
// Keywords: Aspose.Cells C# data validation theme color | Accent5 background Aspose.Cells | list validation Aspose.Cells .NET | apply theme color to cells Aspose | set validation cell style Aspose.Cells
// Common Searches: Aspose.Cells set validation background to theme accent color | C# apply Accent5 theme to validated cells | how to use theme colors in Aspose.Cells styles | list dropdown validation with themed background Aspose | Aspose.Cells example Accent5 background
// Developer Intent: Add a list‑type data validation and highlight the validated cells with the workbook’s Accent5 theme color.
// Use Cases: Insert a Yes/No dropdown in A1:B2 and color the cells with the corporate Accent5 theme for visual consistency. | Generate a template where validation cells inherit the document’s theme, simplifying end‑user editing. | Create reports that automatically apply brand‑aligned theme colors to validation ranges.
// AI Prompts: Write C# code with Aspose.Cells that adds a list validation to A1:B2 and sets the cell background to the workbook’s Accent5 theme color. | Show how to modify the tint of an Accent5 themed background for a validation range in Aspose.Cells. | Provide an Aspose.Cells .NET example that applies a solid pattern style with a theme color to cells containing data validation.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationWithThemeBackground
{
    // Creates a workbook, defines range A1:B2, adds a list‑type validation (Yes/No), applies a solid style whose background uses the workbook’s Accent5 theme color, and saves the file as ValidationWithAccent5Background.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the area (A1:B2) where the validation will be applied
            CellArea validationArea = CellArea.CreateCellArea(0, 0, 1, 1); // rows 0-1, columns 0-1

            // Add a list‑type validation to the worksheet
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.Formula1 = "Yes,No";          // Allowed values
            validation.ShowInput = true;            // Show input message when cell is selected
            validation.InputTitle = "Select Option";
            validation.InputMessage = "Please choose Yes or No.";

            // Create a style that uses the theme's Accent5 color for the background
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;   // Required for background color to be visible
            // Accent5 with no tint (0) – you can adjust the tint value if needed
            style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent5, 0);

            // Apply the style to each cell in the validation area
            for (int row = validationArea.StartRow; row <= validationArea.EndRow; row++)
            {
                for (int col = validationArea.StartColumn; col <= validationArea.EndColumn; col++)
                {
                    cells[row, col].SetStyle(style);
                }
            }

            // Save the workbook
            workbook.Save("ValidationWithAccent5Background.xlsx");
        }
    }
}
