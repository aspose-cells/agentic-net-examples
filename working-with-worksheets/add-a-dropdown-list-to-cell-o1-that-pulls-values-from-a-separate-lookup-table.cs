// Title: Add a drop‑down list to cell O1 that references a lookup table on a separate worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a 'Lookup' worksheet, fills A1:A5 with options, and adds a list‑type data validation to cell O1 on the main sheet referencing that range. | Generate a complete Aspose.Cells program that builds a lookup sheet, applies a dropdown validation to a specific cell, and saves the workbook as an Excel file.
// Common Searches: asp.net cells add data validation list referencing another worksheet c# | c# aspose.cells create dropdown list in cell O1 from lookup range | how to set validation formula to external sheet in Aspose.Cells | example of using a lookup table for Excel data validation with Aspose.Cells | aspnet cells populate lookup sheet and apply list validation
// Tags: Aspose.Cells list validation from another worksheet | C# create Excel dropdown using Aspose.Cells | Aspose.Cells set validation formula to external range | populate lookup sheet for data validation Aspose.Cells | Excel data validation O1 with lookup table Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a new workbook, adds a 'Lookup' worksheet populated with five options, defines a list‑type data validation on cell O1 of the 'Main' worksheet that points to the lookup range, and saves the file as DropDownExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will hold the lookup table
            Worksheet lookupSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            lookupSheet.Name = "Lookup";

            // Populate the lookup table (A1:A5) with sample values
            string[] lookupValues = { "Option1", "Option2", "Option3", "Option4", "Option5" };
            for (int i = 0; i < lookupValues.Length; i++)
            {
                lookupSheet.Cells[i, 0].PutValue(lookupValues[i]); // Column A (index 0)
            }

            // Get the main worksheet where the drop‑down will be placed
            Worksheet mainSheet = workbook.Worksheets[0];
            mainSheet.Name = "Main";

            // Define the cell area (O1) that the validation applies to
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 14, // Column O (0‑based index)
                EndColumn = 14
            };

            // Add a data validation (drop‑down list) to the defined area
            int validationIndex = mainSheet.Validations.Add(area);
            Validation validation = mainSheet.Validations[validationIndex];
            validation.Type = ValidationType.List;                     // List type for drop‑down
            validation.Operator = OperatorType.None;                  // Not needed for list
            validation.Formula1 = "'Lookup'!$A$1:$A$5";                // Reference to lookup range
            validation.ShowError = true;                              // Show error message if invalid
            validation.ErrorTitle = "Invalid Selection";
            validation.ErrorMessage = "Please select a value from the list.";

            // Save the workbook to a file
            workbook.Save("DropDownExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
