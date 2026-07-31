// Title: C# – Merge D2:E2, add dropdown list validation, and save with merged‑area options using Aspose.Cells
// Description: Loads or creates a workbook, merges cells D2:E2 on the first worksheet, applies a list‑type data validation (Option1, Option2, Option3) to the merged range, and saves the file as XLS with MergeAreas and ValidateMergedAreas enabled.
// Keywords: Aspose.Cells | C# | merge cells D2:E2 | dropdown list validation | list validation Aspose.Cells | XlsSaveOptions MergeAreas | ValidateMergedAreas | Excel export merged cells | data validation merged range
// Common Searches: Aspose.Cells merge cells and add dropdown C# | How to apply list validation to a merged range in Aspose.Cells | Save XLS with merged cells and validation using Aspose.Cells | Enable ValidateMergedAreas in XlsSaveOptions | C# code for merged cell dropdown in Aspose.Cells
// Developer Intent: Merge cells D2:E2, attach a dropdown (list) validation to the merged area, and export the workbook while preserving the merged‑cell validation.
// Use Cases: Building a template where a merged header cell offers a predefined selection list. | Generating a legacy XLS report that requires both merged titles and controlled input via dropdowns. | Creating an Excel form where a merged cell consolidates options for user entry and must retain validation after saving.
// AI Prompts: Generate C# code with Aspose.Cells that merges D2:E2, adds a list validation containing custom items, and saves the workbook using XlsSaveOptions with MergeAreas and ValidateMergedAreas set to true. | Explain how the ValidateMergedAreas property influences data validation behavior for merged cells when exporting to XLS with Aspose.Cells. | Show how to assign multiple validation areas to a single merged range in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads or creates a workbook, merges cells D2:E2 on the first worksheet, applies a list‑type data validation (Option1, Option2, Option3) to the merged range, and saves the file as XLS with MergeAreas and ValidateMergedAreas enabled.
class Program
{
    static void Main()
    {
        // Load an existing workbook if it exists; otherwise create a new one
        string inputPath = "input.xlsx";
        Workbook workbook = System.IO.File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells D2:E2 (zero‑based indices: row 1, column 3, 1 row, 2 columns)
        worksheet.Cells.Merge(1, 3, 1, 2);

        // Add a data‑validation list to the merged cells
        Validation validation = worksheet.Validations[worksheet.Validations.Add()];
        validation.Type = ValidationType.List;
        // The list items are provided as a quoted, comma‑separated string
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

        // Save the workbook with merged‑area validation enabled
        XlsSaveOptions saveOptions = new XlsSaveOptions
        {
            MergeAreas = true,
            ValidateMergedAreas = true
        };
        workbook.Save("output.xlsx", saveOptions);
    }
}
