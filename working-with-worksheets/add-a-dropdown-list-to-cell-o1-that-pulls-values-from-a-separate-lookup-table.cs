// Title: C# – Add an In‑Cell Drop‑Down List to O1 Using a Lookup Table with Aspose.Cells
// Description: This example creates a new workbook, fills cells A1:A5 with option values, applies a list‑type data validation to cell O1 that references the A1:A5 range, enables the in‑cell drop‑down, and saves the file as DropDownDemo.xlsx.
// Keywords: Aspose.Cells C# dropdown list | Excel data validation list Aspose.Cells | in‑cell drop‑down Aspose.Cells .NET | lookup range validation Aspose.Cells | C# Excel dropdown example
// Common Searches: Aspose.Cells add dropdown to a cell | C# create Excel list validation from another column | how to use lookup range for data validation Aspose.Cells | in‑cell dropdown example C# Aspose.Cells
// Developer Intent: Add an in‑cell drop‑down list to cell O1 that pulls its items from a lookup table in column A using Aspose.Cells for .NET.
// Use Cases: Design a template where users must select a predefined option in O1, ensuring consistent data entry. | Build a reporting form that restricts input to values maintained in a separate lookup range. | Create a dynamic data‑entry sheet where updating the source range automatically updates the dropdown options.
// AI Prompts: Generate C# code with Aspose.Cells to add a list validation to cell B2 referencing range D1:D10. | Show how to replace the validation formula with a named range in Aspose.Cells. | Suggest robust error‑handling for empty or out‑of‑range lookup tables when creating a dropdown.

using System;
using Aspose.Cells;

// This example creates a new workbook, fills cells A1:A5 with option values, applies a list‑type data validation to cell O1 that references the A1:A5 range, enables the in‑cell drop‑down, and saves the file as DropDownDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a lookup table in column A (A1:A5)
            sheet.Cells["A1"].PutValue("Option1");
            sheet.Cells["A2"].PutValue("Option2");
            sheet.Cells["A3"].PutValue("Option3");
            sheet.Cells["A4"].PutValue("Option4");
            sheet.Cells["A5"].PutValue("Option5");

            // Add data validation (list) to cell O1
            int validationIndex = sheet.Validations.Add();               // create a new validation rule
            Validation validation = sheet.Validations[validationIndex]; // retrieve the rule
            validation.Type = ValidationType.List;                      // list type validation
            validation.Formula1 = "A1:A5";                               // reference the lookup range
            validation.InCellDropDown = true;                           // show in‑cell dropdown
            validation.ShowError = true;                                // optional: display error message

            // Save the workbook
            workbook.Save("DropDownDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
