// Title: C# – Add Whole‑Number Validation (1‑100) to Entire Column C with Aspose.Cells
// Description: Creates a new workbook, defines a CellArea covering all rows of column C, adds a WholeNumber validation with the Between operator (1‑100), sets optional input and error messages, and saves the file as an XLSX document.
// Keywords: Aspose.Cells C# data validation | whole number validation column C | Excel validation 1 to 100 Aspose | restrict cell input Aspose.Cells | C# Aspose.Cells whole number range
// Common Searches: Aspose.Cells add whole number validation to column C | C# set data validation 1-100 in Excel with Aspose | how to limit column C values to integers 1-100 using Aspose.Cells | Aspose.Cells validation input message error alert example
// Developer Intent: Implement a validation rule that forces every cell in column C to accept only whole numbers from 1 to 100.
// Use Cases: Ensure employee ID entries stay within a predefined numeric range. | Prevent inventory quantities from being entered outside the 1‑100 limit. | Validate student scores so only whole numbers between 1 and 100 are allowed.
// AI Prompts: Generate C# code using Aspose.Cells to apply a whole‑number (1‑100) validation to column C with custom input and error messages. | Show how to modify the validation to accept only positive integers greater than zero in column C. | Explain how to copy the same whole‑number validation to multiple columns in a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Creates a new workbook, defines a CellArea covering all rows of column C, adds a WholeNumber validation with the Between operator (1‑100), sets optional input and error messages, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a validation area that covers the entire column C (index 2)
            // Excel has 1,048,576 rows (0‑based index 0 to 1,048,575)
            CellArea columnCArea = CellArea.CreateCellArea(0, 2, 1048575, 2);

            // Add a new validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(columnCArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation: whole numbers between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Optional: display input message and error alert
            validation.ShowInput = true;
            validation.InputTitle = "Enter Number";
            validation.InputMessage = "Please enter a whole number between 1 and 100.";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "The value must be a whole number between 1 and 100.";
            validation.AlertStyle = ValidationAlertType.Stop;

            // Save the workbook to a file
            workbook.Save("ColumnC_WholeNumber_1_to_100.xlsx");
        }
    }
}
