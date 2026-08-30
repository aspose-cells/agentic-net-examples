// Title: Create a list‑type data validation dropdown for column S (S1:S100) in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook and adds a list‑type validation to cells S1 through S100, allowing only Apple, Orange, Banana, and Cherry, with an in‑cell dropdown. | Write C# using Aspose.Cells to define a validation list whose allowed values are read from a named range on another worksheet instead of a hard‑coded string. | Produce a C# sample that applies the same list validation to multiple columns (e.g., S, T, U) and sets custom input and error messages for each column.
// Common Searches: Aspose.Cells C# how to add a drop‑down list validation to column S | C# set data validation list for Excel range S1:S100 using Aspose.Cells | restrict Excel cell values to a predefined list with Aspose.Cells .NET | apply list validation to multiple columns in Aspose.Cells example | load validation list from another worksheet in Aspose.Cells C#
// Tags: Aspose.Cells list validation Excel column | Aspose.Cells data validation dropdown .NET | Excel column S validation Aspose.Cells | Aspose.Cells validation formula list C# | Aspose.Cells apply validation to range

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // The example creates a new workbook, defines a validation area covering cells S1 to S100, adds a list‑type validation with the allowed values Apple, Orange, Banana, and Cherry, shows an in‑cell dropdown, customizes input and error messages, and saves the file as ColumnS_Validation.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the validation area for column S (index 18)
            // Here we apply the validation to rows 0 through 99 (i.e., S1:S100)
            CellArea validationArea = CellArea.CreateCellArea(0, 18, 99, 18);

            // Add a new validation to the worksheet for the defined area
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];

            // Set the validation type to List and provide the allowed values
            validation.Type = ValidationType.List;
            validation.Formula1 = "Apple,Orange,Banana,Cherry"; // predefined array of strings
            validation.InCellDropDown = true; // show dropdown in the cell

            // Optional: customize the input and error messages
            validation.InputTitle = "Select a Fruit";
            validation.InputMessage = "Choose one of the listed fruits.";
            validation.ErrorTitle = "Invalid Selection";
            validation.ErrorMessage = "Please select a valid fruit from the list.";
            validation.AlertStyle = ValidationAlertType.Stop;
            validation.ShowError = true;
            validation.ShowInput = true;

            // Save the workbook to a file
            workbook.Save("ColumnS_Validation.xlsx");
        }
    }
}
