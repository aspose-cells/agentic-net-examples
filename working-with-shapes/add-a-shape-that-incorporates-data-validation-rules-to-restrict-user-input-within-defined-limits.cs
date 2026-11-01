// Title: Create a rectangle shape over cell B3 and enforce integer (1‑100) data validation with custom prompts using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape at B3, applies whole‑number validation between 1 and 100 to the underlying cell, and sets custom error and input messages with Aspose.Cells. | Write a program that creates an Excel workbook, places a shape containing a prompt, and restricts user entry to integers 1‑100, displaying a validation dialog on invalid input. | Produce a C# snippet that links a rectangle shape to a specific cell, configures the cell’s validation to accept only numbers within a defined range, and saves the workbook.
// Common Searches: aspnet cells add rectangle shape with data validation to a specific cell | c# Aspose.Cells set whole number validation range 1 to 100 with custom error message | how to display input prompt for cell validation using Aspose.Cells | link shape to cell B3 and enforce integer range in Excel via Aspose.Cells
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells whole number validation | Aspose.Cells custom validation error message | Aspose.Cells input prompt for cell | Aspose.Cells shape linked to cell

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a rectangle shape at cell B3, applies whole‑number data validation (1‑100) with custom error and input prompts to that cell, and saves the file as ShapeWithValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a rectangle shape at cell B3 (row 2, column 2) with specified size
            var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 50);
            shape.Text = "Enter value (1‑100)";

            // Define the target cell (B3) and set an initial placeholder value
            var targetCell = sheet.Cells["B3"];
            targetCell.PutValue(0);

            // Define the cell area for validation (B3)
            var area = new CellArea
            {
                StartRow = 2,
                StartColumn = 1,
                EndRow = 2,
                EndColumn = 1
            };

            // Add whole number validation (integer between 1 and 100)
            int validationIndex = sheet.Validations.Add(area);
            var validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;   // Validate as whole number
            validation.Operator = OperatorType.Between;    // Value must be between two limits
            validation.Formula1 = "1";                      // Lower limit
            validation.Formula2 = "100";                    // Upper limit
            validation.ShowError = true;                    // Show error dialog on invalid entry
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please enter an integer between 1 and 100.";
            // Input prompt (shown when the cell is selected)
            validation.InputTitle = "Input Required";
            validation.InputMessage = "Enter a number from 1 to 100.";

            // Save the workbook
            workbook.Save("ShapeWithValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
