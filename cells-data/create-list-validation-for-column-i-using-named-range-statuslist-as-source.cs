// Title: Create Excel list validation for column I using the named range 'StatusList' with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add data validation list to entire column I from a named range | How to set a dropdown list in column I using a named range with Aspose.Cells .NET | C# Aspose.Cells create list validation with formula =StatusList for column I | Apply Excel data validation to a whole column using Aspose.Cells API in C# | Use ValidationType.List with a named range for column I in Aspose.Cells .NET
// Tags: named range source list validation Aspose.Cells | apply data validation to full column C# | column I dropdown list Aspose.Cells | ValidationType.List formula named range .NET | create Excel validation using Aspose.Cells API

using Aspose.Cells;
using System;

// The example creates a new workbook, defines the entire column I as a CellArea, adds a list‑type validation whose source is the named range 'StatusList', enables an in‑cell dropdown, sets input and error messages, and saves the file as ListValidation.xlsx.
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

            // Define the cell area for column I (entire column)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 1048575,   // maximum rows in Excel
                StartColumn = 8,    // column I (zero‑based index)
                EndColumn = 8
            };

            // Add list validation to the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            // Use the named range 'StatusList' as the source list
            validation.Formula1 = "=StatusList";
            // Enable dropdown in the cell
            validation.InCellDropDown = true;

            // Optional: user prompts (messages are shown automatically when set)
            validation.InputMessage = "Select a status";
            validation.ErrorMessage = "Invalid status selected";

            // Save the workbook
            workbook.Save("ListValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
