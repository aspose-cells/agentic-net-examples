// Title: Create a drop‑down list validation for cell U2 with Aspose.Cells for .NET
// Description: Shows how to generate a workbook, target cell U2, add a list‑type data validation with an in‑cell drop‑down, set the allowed values (OptionA, OptionB, OptionC) via Formula1, and save the file as U2Validation.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# data validation | list validation | drop‑down list | cell U2 | ValidationType.List | Formula1 | .NET Excel | predefined options | spreadsheet dropdown
// Common Searches: Aspose.Cells create drop‑down list for a single cell | C# add list validation to cell U2 using Aspose.Cells | How to set ValidationType.List with Formula1 in Aspose.Cells | Excel data validation with predefined options in .NET | Aspose.Cells restrict cell input to specific values
// Developer Intent: Add a data‑validation rule that limits cell U2 to a predefined list of values.
// Use Cases: Force users to choose a category from a fixed list in a generated report. | Provide a status selector (e.g., OptionA‑OptionC) in a data‑entry template. | Prevent invalid entries in a form by applying a list validation to cell U2.
// AI Prompts: Write C# code with Aspose.Cells that adds a list validation to cell U2 containing 'OptionA', 'OptionB', and 'OptionC' and saves the workbook. | Explain how to change the validation to reference a named range instead of a hard‑coded list in Aspose.Cells. | Show how to apply the same list validation to a range such as U2:U10 using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to generate a workbook, target cell U2, add a list‑type data validation with an in‑cell drop‑down, set the allowed values (OptionA, OptionB, OptionC) via Formula1, and save the file as U2Validation.xlsx using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell area for U2 (zero‑based indices: row 1, column 20)
            CellArea u2Area = new CellArea
            {
                StartRow = 1,
                StartColumn = 20,
                EndRow = 1,
                EndColumn = 20
            };

            // Add a validation rule for the specified area
            int validationIndex = worksheet.Validations.Add(u2Area);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation as a list with allowed options
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            // Formula1 must be quoted when using a comma‑separated list
            validation.Formula1 = "\"OptionA,OptionB,OptionC\"";

            // Save the workbook
            string outputPath = "U2Validation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
