// Title: Add a list‑type data validation dropdown to column A and freeze the column using Aspose.Cells for .NET (C#)
// Common Searches: how to add a list validation dropdown to a column in Excel using Aspose.Cells C# | Aspose.Cells C# freeze first column after setting data validation | C# programmatically apply list data validation to A2:A100 and keep it visible while scrolling | using Aspose.Cells to create dropdown list validation and freeze panes in an Excel workbook | example of FreezePanes with data validation in Aspose.Cells for .NET
// Tags: list validation dropdown Aspose.Cells | freeze first column Aspose.Cells | apply validation to A2:A100 C# | Excel workbook with frozen pane and validation | Aspose.Cells data validation example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, adds a list‑type data validation with an in‑cell dropdown to cells A2 through A100, sets input and error messages, freezes column A using FreezePanes so the validation stays visible while scrolling, and saves the file as ValidatedAndFrozen.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for validation (A2:A100)
            CellArea validationArea = new CellArea
            {
                StartRow = 1,   // Row 2 (0‑based index)
                EndRow = 99,    // Row 100
                StartColumn = 0,
                EndColumn = 0
            };

            // Add validation for the specified range – Add returns the index of the new validation
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to List with in‑cell dropdown
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Configure input (prompt) message
            validation.InputMessage = "Select a value from the list.";
            validation.InputTitle = "Input Required";
            // ShowInputMessage property is not required; default behavior displays the message

            // Configure error message
            validation.ErrorMessage = "Please select a valid option.";
            validation.ErrorTitle = "Invalid Input";
            validation.ShowError = true;

            // Freeze the first column (column A) so the validation remains visible while scrolling
            // Parameters: rowSplit, columnSplit, topRow, leftColumn
            sheet.FreezePanes(0, 1, 0, 1);

            // Determine output path and save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ValidatedAndFrozen.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
