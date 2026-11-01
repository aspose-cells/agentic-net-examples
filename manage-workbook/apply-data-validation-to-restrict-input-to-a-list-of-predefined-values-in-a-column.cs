// Title: Add a list‑type data validation with drop‑down to column A (rows 2‑100) using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to create a list data validation for a column range | C# add in‑cell drop‑down list validation to Excel column using Aspose.Cells | Set predefined values for cells in column A rows 2 to 100 with Aspose.Cells | Configure input and error messages for list validation in Aspose.Cells C# | Apply data validation list with Apple, Banana, Cherry using Aspose.Cells .NET
// Tags: list data validation Aspose.Cells .NET | in‑cell dropdown Aspose.Cells C# | validation range CellArea Aspose.Cells | custom input message Aspose.Cells validation | error message configuration Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, defines a CellArea covering column A rows 2‑100, adds a list‑type validation with the values Apple, Banana, and Cherry, enables the in‑cell drop‑down, sets custom input and error messages, and saves the workbook as DataValidationExample.xlsx.
class DataValidationExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the validation (column A, rows 2 to 100).
            CellArea validationRange = new CellArea
            {
                StartRow = 1,      // Row index is zero‑based (1 = second row)
                StartColumn = 0,   // Column A
                EndRow = 99,       // Row 100
                EndColumn = 0      // Column A
            };

            // Add a new validation for the defined range and retrieve the Validation object.
            int validationIndex = sheet.Validations.Add(validationRange);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to List and provide allowed values.
            validation.Type = ValidationType.List;
            validation.Operator = OperatorType.Equal;
            validation.Formula1 = "\"Apple,Banana,Cherry\"";

            // Enable the drop‑down list in the cells.
            validation.InCellDropDown = true;

            // Configure input (prompt) message.
            validation.InputTitle = "Select a fruit";
            validation.InputMessage = "Please choose a fruit from the list.";

            // Configure error message.
            validation.ErrorTitle = "Invalid entry";
            validation.ErrorMessage = "The value must be one of the predefined options.";

            // Save the workbook.
            workbook.Save("DataValidationExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
