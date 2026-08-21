// Title: Apply list‑type data validation to column S (rows 0‑99) using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, defines the range S1:S100, adds a List validation with the allowed values "Apple", "Banana" and "Cherry", configures a stop‑alert with a custom error title and message, and saves the file as ColumnSValidation.xlsx.
// Keywords: Aspose.Cells C# data validation | list validation column S | Excel drop‑down list Aspose | restrict cell values .NET | Aspose.Cells validation example | Excel data validation programmatically | C# Aspose.Cells workbook validation
// Common Searches: Aspose.Cells add drop‑down list to column S | C# restrict column values to specific strings in Excel | How to create list validation with Aspose.Cells for .NET | Set error message for data validation in Aspose.Cells | Excel data validation using Aspose.Cells C#
// Developer Intent: Add a List‑type validation to column S (rows 0‑99) that only permits predefined string values.
// Use Cases: Standardize fruit selection in a generated report by providing a drop‑down list in column S. | Prevent invalid text entry in a critical data column before downstream processing. | Enforce consistent terminology across exported Excel files created with Aspose.Cells.
// AI Prompts: Generate C# code with Aspose.Cells to apply a list validation to column S for rows 0‑99 using the values "Apple", "Banana", "Cherry" and display a stop‑alert on invalid input. | Explain how to reference a dynamic named range on another worksheet for list validation in Aspose.Cells. | Provide a step‑by‑step guide to add different list validations to multiple columns within the same workbook using Aspose.Cells for .NET.

using Aspose.Cells;
using System;

// Creates a new workbook, defines the range S1:S100, adds a List validation with the allowed values "Apple", "Banana" and "Cherry", configures a stop‑alert with a custom error title and message, and saves the file as ColumnSValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation area for column S (index 18) rows 0 to 99
        CellArea area = CellArea.CreateCellArea(0, 18, 99, 18);

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Set validation type to List and provide the allowed string values
        validation.Type = ValidationType.List;
        validation.InCellDropDown = true;
        validation.Value1 = new object[] { "Apple", "Banana", "Cherry" };

        // Configure alert style and error messages (optional)
        validation.AlertStyle = ValidationAlertType.Stop;
        validation.ErrorTitle = "Invalid Entry";
        validation.ErrorMessage = "Please select a value from the list.";
        validation.ShowError = true;

        // Save the workbook
        workbook.Save("ColumnSValidation.xlsx");
    }
}
