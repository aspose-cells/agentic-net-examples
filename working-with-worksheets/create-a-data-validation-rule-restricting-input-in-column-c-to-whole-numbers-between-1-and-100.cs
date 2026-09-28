// Title: Add whole-number (1‑100) data validation to the entire column C using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that applies a whole‑number validation between 1 and 100 to every cell in column C of a worksheet. | Show how to create a CellArea covering column C and configure a Validation object with Type=WholeNumber and Operator=Between in Aspose.Cells. | Demonstrate setting a custom error title and message for a numeric range validation on column C using Aspose.Cells in C#.
// Common Searches: aspnet aspose.cells column C numeric range validation 1 to 100 | c# restrict Excel column values to a number between 1 and 100 with Aspose.Cells | how to apply a Between operator validation to an entire column using Aspose.Cells .NET | set custom error title and message for data validation in Aspose.Cells worksheet | aspose.cells example for numeric validation in C#
// Tags: Aspose.Cells numeric range validation C# | column C validation Aspose.Cells | Validation Type WholeNumber Aspose.Cells | Operator Between Aspose.Cells | custom error title Aspose.Cells validation

using Aspose.Cells;
using System;

// The sample creates a new workbook, defines a CellArea that spans the whole of column C, adds a whole‑number validation with a Between operator limited to 1‑100, configures optional error title and message, and saves the file as DataValidation.xlsx.
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

            // Define the range for the entire column C (C1:C1048576)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 1048575, // last possible row (zero‑based index)
                StartColumn = 2, // column C (zero‑based index)
                EndColumn = 2
            };

            // Add a validation rule to the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Restrict to whole numbers between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Optional UI settings
            validation.IgnoreBlank = true;
            validation.InCellDropDown = false;
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please enter a whole number between 1 and 100.";

            // Save the workbook
            workbook.Save("DataValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
