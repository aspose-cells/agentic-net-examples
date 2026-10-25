// Title: Create a date‑validation rule that only allows dates after today in cell M2 with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set data validation for future dates in a specific cell | How to restrict Excel cell M2 to dates later than today using Aspose.Cells | Add date greater than TODAY() validation to cell M2 in a .NET workbook | C# Aspose.Cells validation operator GreaterThan with formula TODAY() example | Apply input and error messages for date validation in Aspose.Cells
// Tags: Aspose.Cells date validation after today | C# set validation on cell M2 with Aspose.Cells | ValidationType.Date with OperatorType.GreaterThan example | Aspose.Cells custom error title and message for validation | Saving workbook after applying cell validation in .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, defines cell M2 using a CellArea, adds a date validation rule that requires the entered date to be greater than TODAY(), configures custom error and input prompts, and saves the workbook as Output.xlsx, all implemented in C# with Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell area for M2 (column 12, row 1 in zero‑based indexing)
            CellArea area = new CellArea
            {
                StartRow = 1,
                EndRow = 1,
                StartColumn = 12,
                EndColumn = 12
            };

            // Add a data validation rule for the defined area
            ValidationCollection validations = sheet.Validations;
            int validationIndex = validations.Add(area); // use the overload that accepts CellArea
            Validation validation = validations[validationIndex];

            // Configure validation: date greater than TODAY()
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.GreaterThan;
            validation.Formula1 = "=TODAY()";

            // Configure error message
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "Please enter a date after today.";

            // Configure input (prompt) message
            validation.ShowInput = true;
            validation.InputTitle = "Date Required";
            validation.InputMessage = "Enter a date later than today.";

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
