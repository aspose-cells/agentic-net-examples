// Title: Create a column‑wise drop‑down list using Aspose.Cells for .NET (C#)
// Description: The sample builds a new workbook, selects the range B2:B20, adds a List‑type data validation with the values "OptionA,OptionB,OptionC", enables the in‑cell drop‑down, shows a custom input prompt, and saves the file as ColumnListValidation.xlsx.
// Keywords: Aspose.Cells | C# data validation | .NET Excel drop‑down list | list validation | in‑cell dropdown | predefined values | column validation | CellArea | ValidationType.List | input message | Excel workbook automation
// Common Searches: Aspose.Cells add list validation C# | C# create drop‑down list in Excel with Aspose | How to restrict column values using Aspose.Cells | Set input message for data validation Aspose.Cells | Create Excel data validation list .NET
// Developer Intent: The developer needs to limit entries in a specific column to a set of predefined options by applying a drop‑down list validation.
// Use Cases: Design a template where users pick a status from a drop‑down in column B rows 2‑20. | Enforce allowed categories in a data‑entry sheet by applying a list validation to a defined range. | Provide an input prompt that guides users to select one of the permitted values when editing cells.
// AI Prompts: Generate C# code with Aspose.Cells that applies a list validation to a dynamic range and includes a custom input message. | Show how to modify the validation to reference a named range instead of hard‑coded options. | Explain how to read the selected value from a validated cell after the workbook is saved.

using System;
using Aspose.Cells;

// The sample builds a new workbook, selects the range B2:B20, adds a List‑type data validation with the values "OptionA,OptionB,OptionC", enables the in‑cell drop‑down, shows a custom input prompt, and saves the file as ColumnListValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the column range (e.g., B2:B20) where validation will be applied
        // Row and column indices are zero‑based
        CellArea area = CellArea.CreateCellArea(1, 1, 19, 1); // rows 2‑20, column B

        // Add a validation for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation as a list with predefined values
        validation.Type = ValidationType.List;
        validation.Formula1 = "OptionA,OptionB,OptionC";

        // Enable the in‑cell drop‑down list
        validation.InCellDropDown = true;

        // Optional: display an input message when the cell is selected
        validation.ShowInput = true;
        validation.InputTitle = "Select Value";
        validation.InputMessage = "Choose one of the predefined options.";

        // Save the workbook
        workbook.Save("ColumnListValidation.xlsx");
    }
}
