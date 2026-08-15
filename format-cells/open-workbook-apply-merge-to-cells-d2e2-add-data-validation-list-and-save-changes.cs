// Title: C# Aspose.Cells: Merge D2:E2 and Add a Dropdown List Validation
// Description: Loads an existing Excel workbook, merges cells D2:E2 on the first worksheet, creates a list‑type data validation (e.g., "Option1,Option2,Option3"), applies the validation to the merged range, and saves the updated file.
// Keywords: Aspose.Cells | C# | .NET | merge cells | Excel merge D2:E2 | data validation list | dropdown list | merged cells validation | save workbook | list validation
// Common Searches: Aspose.Cells merge cells D2 E2 C# | add dropdown validation to merged cells Aspose.Cells | C# example list validation for merged range Excel | how to apply data validation to merged cells using Aspose.Cells | save workbook after adding validation Aspose.Cells .NET
// Developer Intent: Merge cells D2:E2, attach a list‑type data validation, and save the workbook.
// Use Cases: Create a header that spans D2:E2 with a predefined dropdown for user selection. | Build a template where merged title cells enforce entry from a specific list. | Generate reports that require consistent values in merged cells via a dropdown.
// AI Prompts: Write C# code with Aspose.Cells to merge D2:E2, add a list validation containing custom items, and save the workbook. | Show how to replace the hard‑coded validation list with a reference to a named range in Aspose.Cells. | Provide robust error handling for missing input files and invalid validation formulas when merging cells and adding data validation.

using System;
using Aspose.Cells;

// Loads an existing Excel workbook, merges cells D2:E2 on the first worksheet, creates a list‑type data validation (e.g., "Option1,Option2,Option3"), applies the validation to the merged range, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells D2:E2 (zero‑based row 1, column 3, 1 row, 2 columns)
        worksheet.Cells.Merge(1, 3, 1, 2);

        // Add a data‑validation list to the merged cells
        Validation validation = worksheet.Validations[worksheet.Validations.Add()];
        validation.Type = ValidationType.List;
        // List items are provided as a comma‑separated string enclosed in quotes
        validation.Formula1 = "\"Option1,Option2,Option3\"";

        // Define the cell area that the validation applies to (D2:E2)
        CellArea area = new CellArea
        {
            StartRow = 1,
            StartColumn = 3,
            EndRow = 1,
            EndColumn = 4
        };
        validation.AddArea(area);

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
