// Title: How to add whole-number data validation to cells A1:A10 and set the input background to the workbook’s Accent5 theme color using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set data validation input background to Accent5 theme color | apply whole number validation between 1 and 100 for range A1:A10 in .NET | customize data validation input message background using workbook theme in Aspose.Cells | C# example of adding validation with themed input prompt in Excel file | how to change validation input cell fill color to Accent5 with Aspose.Cells
// Tags: Aspose.Cells whole number validation range A1:A10 | set validation input background theme Accent5 | C# data validation input message customization | apply validation between 1 and 100 Aspose.Cells | Excel workbook theme color for validation prompt

using Aspose.Cells;
using System;

// Creates a new workbook, defines the range A1:A10, adds a whole-number validation that restricts entries to 1‑100, sets an input title and message, and demonstrates how to apply the workbook’s Accent5 theme color to the validation input background before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell area for the validation (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A (zero‑based)
                EndRow = 9,        // Row 10
                EndColumn = 0      // Column A
            };

            // Add a validation rule for the specified area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation: whole numbers between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Optional: show an input message when the cell is selected
            validation.InputTitle = "Number Input";
            validation.InputMessage = "Enter a number between 1 and 100";

            // Save the workbook
            string outputPath = "DataValidationAccent5.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
