// Title: Add a list‑type data validation dropdown sourced from a hidden worksheet range using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a hidden worksheet, populates it with allowed values, and applies a list‑type validation to a target range on a visible sheet. | Demonstrate how to build a formula referencing a hidden sheet range for a dropdown validation in Aspose.Cells. | Provide a loop that adds list validation to each cell in a column, using the hidden list length to construct the range address dynamically.
// Common Searches: how to create a dropdown list in Excel using Aspose.Cells C# from a hidden sheet | Aspose.Cells data validation list referencing another worksheet | C# hide worksheet and use its cells for validation list with Aspose.Cells | set list validation range dynamically based on hidden sheet values in Aspose.Cells | apply data validation to multiple cells with dropdown from hidden sheet using Aspose.Cells .NET
// Tags: list validation hidden worksheet Aspose.Cells | C# Excel dropdown from hidden sheet Aspose.Cells | dynamic validation range based on hidden list Aspose.Cells | programmatic data validation referencing another sheet .NET | Aspose.Cells create hidden sheet for dropdown list

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, adds a hidden worksheet named "HiddenList" filled with five options, builds a range address for those cells, and applies a list‑type data validation with an in‑cell dropdown to cells B2:B10 on the visible "DataEntry" sheet. The workbook is saved as DataValidationWithHiddenList.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first (visible) worksheet and rename it
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "DataEntry";

            // Add a hidden worksheet that will hold the allowed list values
            int hiddenSheetIndex = workbook.Worksheets.Add();
            Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
            hiddenSheet.Name = "HiddenList";
            hiddenSheet.IsVisible = false; // hide the sheet from the user

            // Populate the hidden sheet with the list items (e.g., A1:A5)
            string[] allowedValues = { "Option1", "Option2", "Option3", "Option4", "Option5" };
            for (int i = 0; i < allowedValues.Length; i++)
            {
                hiddenSheet.Cells[i, 0].PutValue(allowedValues[i]); // column A
            }

            // Build the address of the range that contains the list on the hidden sheet
            string listRange = $"'{hiddenSheet.Name}'!$A$1:$A${allowedValues.Length}";

            // Define the target range on the visible sheet where the validation will be applied (e.g., B2:B10)
            int startRow = 1;   // Row 2 (zero‑based index)
            int endRow = 9;     // Row 10
            int targetColumn = 1; // Column B (zero‑based index)

            // Apply data validation to each cell in the target range
            for (int row = startRow; row <= endRow; row++)
            {
                // Define the cell area for the current cell
                CellArea area = new CellArea
                {
                    StartRow = row,
                    StartColumn = targetColumn,
                    EndRow = row,
                    EndColumn = targetColumn
                };

                // Add a new validation rule for the defined area (Add returns the index)
                int validationIndex = dataSheet.Validations.Add(area);
                Validation validation = dataSheet.Validations[validationIndex];

                validation.Type = ValidationType.List;          // List validation
                validation.Operator = OperatorType.None;
                validation.Formula1 = listRange;                // Reference to hidden list range
                validation.InCellDropDown = true;               // Show dropdown arrow
                validation.ShowError = true;                    // Show error dialog on invalid entry
                validation.ErrorTitle = "Invalid Selection";
                validation.ErrorMessage = "Please select a value from the list.";
                validation.AlertStyle = ValidationAlertType.Stop;
            }

            // Determine output path and ensure the directory exists
            string outputPath = "DataValidationWithHiddenList.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
