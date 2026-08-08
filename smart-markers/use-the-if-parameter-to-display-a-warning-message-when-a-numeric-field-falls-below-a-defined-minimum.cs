// Title: Aspose.Cells C# – Show a warning when a numeric cell is below a defined minimum
// Description: Creates an Excel workbook, inserts sample numbers, and adds a whole‑number validation that uses the LessThan operator. The rule is set to the Warning alert style with a custom title and message, so any entry lower than the specified minimum triggers a non‑blocking warning before the file is saved.
// Keywords: Aspose.Cells | C# | Excel data validation | warning alert | LessThan operator | minimum numeric value | smart markers example | US developers | EU developers
// Common Searches: Aspose.Cells show warning for low numeric value C# | Excel validation less than threshold with warning dialog using Aspose.Cells | C# data validation warning style Aspose.Cells example | How to add a warning alert to a cell range in Aspose.Cells
// Developer Intent: Add a validation rule that displays a warning message when a cell’s numeric entry is smaller than a preset minimum.
// Use Cases: Enforce a minimum purchase amount in a sales ledger while allowing the user to continue after acknowledging the warning. | Alert inventory managers when a stock count is entered below the reorder level without blocking the spreadsheet. | Provide immediate feedback for age fields in a registration form, warning when the age is below the legal threshold.
// AI Prompts: Generate Aspose.Cells C# code that warns if values in B2:B15 are less than 5, using a custom error title. | Explain how to replace the constant minimum with a cell reference or named range in the validation rule. | Show how to apply the same warning validation to an entire column and customize the message based on the column header.

using System;
using Aspose.Cells;

namespace AsposeCellsIfWarningDemo
{
    // Creates an Excel workbook, inserts sample numbers, and adds a whole‑number validation that uses the LessThan operator. The rule is set to the Warning alert style with a custom title and message, so any entry lower than the specified minimum triggers a non‑blocking warning before the file is saved.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the minimum allowed value
            const int minAllowed = 10;

            // Add some sample data (some values below the minimum)
            sheet.Cells["A1"].PutValue(5);   // Below minimum – should trigger warning
            sheet.Cells["A2"].PutValue(12);  // Above minimum – no warning
            sheet.Cells["A3"].PutValue(8);   // Below minimum – should trigger warning

            // Create a validation rule for the range A1:A3
            Validation validation = sheet.Validations[sheet.Validations.Add()];
            validation.AddArea(CellArea.CreateCellArea("A1", "A3"));

            // Set validation to WholeNumber and use the "LessThan" operator
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.LessThan;
            validation.Formula1 = minAllowed.ToString(); // Minimum threshold

            // Configure the warning style (not a stop error)
            validation.AlertStyle = ValidationAlertType.Warning; // Show a warning dialog
            validation.ShowError = true;                         // Enable the message
            validation.ErrorTitle = "Value Too Low";
            validation.ErrorMessage = $"The entered number must be greater than or equal to {minAllowed}.";

            // Save the workbook
            workbook.Save("IfWarningDemo.xlsx");
        }
    }
}
