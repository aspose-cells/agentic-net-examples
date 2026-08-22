// Title: Enforce non‑blank entries in column V with a custom stop‑alert using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that adds a required‑field validation to column V (V1:V1000) in an Aspose.Cells workbook and displays a stop‑style error dialog when the cell is left empty. | Write a C# snippet that configures Aspose.Cells to reject blank values in column V and shows a custom error title and message.
// Common Searches: how to add required field validation to column V in Excel with Aspose.Cells C# | Aspose.Cells ignoreblank false example for specific column | C# Aspose.Cells custom error alert for empty cells in worksheet | prevent blank entries in a column using Aspose.Cells validation | set stop style validation message for column V in Aspose.Cells .NET
// Tags: Aspose.Cells required field validation column V | C# Aspose.Cells set IgnoreBlank false | Aspose.Cells custom stop error alert | Excel column V data validation using Aspose.Cells | Aspose.Cells validation error message customization

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Creates a new workbook, defines a validation range covering column V (rows 1‑1001), sets the validation type to AnyValue with IgnoreBlank = false, configures a stop‑style error alert with custom title and message, adds an input prompt, and saves the file as ColumnV_NoBlank_Validation.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the validation area for column V (index 21) from row 0 to row 1000
            CellArea validationArea = CellArea.CreateCellArea(0, 21, 1000, 21);

            // Add a new validation to the worksheet for the defined area
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];

            // Set validation type to AnyValue (allows any non‑blank value)
            validation.Type = ValidationType.AnyValue;

            // Do NOT ignore blank cells – this will disallow blank entries
            validation.IgnoreBlank = false;

            // Configure the error alert that will be shown when a blank is entered
            validation.AlertStyle = ValidationAlertType.Stop;      // Show a stop alert
            validation.ErrorTitle = "Invalid Input";              // Title of the alert dialog
            validation.ErrorMessage = "Blank entries are not allowed in column V."; // Message shown to the user
            validation.ShowError = true;                          // Ensure the error message is displayed

            // Optionally, show an input message when the cell is selected
            validation.ShowInput = true;
            validation.InputTitle = "Required Field";
            validation.InputMessage = "Please enter a value in this cell.";

            // Save the workbook to a file
            workbook.Save("ColumnV_NoBlank_Validation.xlsx");
        }
    }
}
