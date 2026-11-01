// Title: Create a whole‑number data validation (1‑100) for the entire column C in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that adds a validation rule restricting column C to integers from 1 to 100, includes an input prompt and error alert, and saves the workbook. | Write a .NET console example that defines a CellArea covering C:C, sets Validation.Type to WholeNumber, Operator to Between, and configures Formula1/Formula2 for 1 and 100 using Aspose.Cells. | Provide an Aspose.Cells C# snippet that enforces a numeric range of 1 to 100 on every cell in column C, displays a custom input message, and saves the file as DataValidation.xlsx.
// Common Searches: aspnet how to apply whole number validation to C‑column with Aspose.Cells | c# Aspose.Cells data validation for entire C‑column range | set input message and error handling for numeric validation in Excel using Aspose.Cells | restrict Excel C‑column values between 1 and 100 using Aspose.Cells C# example
// Tags: Aspose.Cells integer constraint for specific column | C# define validation range across entire column Aspose | Excel numeric range rule via Aspose.Cells API | Aspose.Cells custom input title and error handling | Apply data validation to full column using Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, defines a validation area covering the whole of column C, adds a whole‑number rule that only permits values from 1 to 100, configures a custom input prompt and error alert, and saves the file as DataValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column C (C:C) – rows 0 to 1,048,575 (Excel's max rows)
            int firstRow = 0;
            int lastRow = 1048575; // Excel 2007+ max rows
            int columnC = 2; // Zero‑based index: 0=A, 1=B, 2=C

            // Build the cell area for the validation range
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = columnC,
                EndRow = lastRow,
                EndColumn = columnC
            };

            // Add a data validation rule to the specified range
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Restrict input to whole numbers between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Optional: show an input prompt
            validation.ShowInput = true;
            validation.InputTitle = "Enter a number";
            validation.InputMessage = "Please enter a whole number between 1 and 100.";

            // Optional: show an error alert when the rule is violated
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid input";
            validation.ErrorMessage = "The value must be a whole number between 1 and 100.";

            // Save the workbook
            workbook.Save("DataValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
