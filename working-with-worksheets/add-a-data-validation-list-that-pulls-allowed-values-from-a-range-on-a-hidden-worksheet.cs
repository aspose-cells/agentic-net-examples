// Title: Aspose.Cells C# – Create a drop‑down data validation list using a hidden worksheet range
// Description: Shows how to build an Excel file with Aspose.Cells, add a hidden sheet that stores lookup values, and attach a List‑type data validation to cell B2 on the main sheet that references the hidden range (e.g., =HiddenData!$A$1:$A$5). The validation displays an in‑cell drop‑down and the workbook is saved as DataValidationFromHiddenSheet.xlsx.
// Keywords: Aspose.Cells | C# | data validation list | hidden worksheet | drop‑down list | Excel validation | ValidationType.List | Formula1 reference | .NET Excel automation
// Common Searches: Aspose.Cells add data validation list from hidden sheet | C# drop‑down list using hidden worksheet range Aspose.Cells | reference hidden sheet in data validation Aspose.Cells | how to hide source list for Excel validation with Aspose.Cells | Aspose.Cells create hidden lookup sheet for validation
// Developer Intent: Add a drop‑down list to a cell that pulls its allowed values from a range on a hidden worksheet using Aspose.Cells for .NET.
// Use Cases: Keep lookup tables out of the user view while still providing selectable options. | Reuse a single hidden list across multiple worksheets without duplicating data. | Protect source data by hiding the sheet yet allowing end‑users to choose values via an in‑cell drop‑down.
// AI Prompts: Generate C# code with Aspose.Cells that creates a hidden sheet, fills it with values, and applies a List validation to a range of cells referencing that hidden sheet. | Explain how to convert the static hidden range into a dynamic named range for data validation in Aspose.Cells. | Show how to copy the hidden‑sheet validation list to several cells or to other worksheets programmatically.

using System;
using Aspose.Cells;

// Shows how to build an Excel file with Aspose.Cells, add a hidden sheet that stores lookup values, and attach a List‑type data validation to cell B2 on the main sheet that references the hidden range (e.g., =HiddenData!$A$1:$A$5). The validation displays an in‑cell drop‑down and the workbook is saved as DataValidationFromHiddenSheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a hidden worksheet to store the allowed values
        int hiddenIndex = workbook.Worksheets.Add();
        Worksheet hiddenSheet = workbook.Worksheets[hiddenIndex];
        hiddenSheet.Name = "HiddenData";
        hiddenSheet.IsVisible = false; // Hide the sheet

        // Populate the hidden sheet with the list items (A1:A5)
        string[] allowedValues = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
        for (int i = 0; i < allowedValues.Length; i++)
        {
            hiddenSheet.Cells[i, 0].PutValue(allowedValues[i]); // Column A
        }

        // Get the main worksheet where the validation will be applied
        Worksheet mainSheet = workbook.Worksheets[0];
        mainSheet.Name = "Main";

        // Define the cell (e.g., B2) that will have the drop‑down list
        CellArea validationArea = CellArea.CreateCellArea(1, 1, 1, 1); // Row 2, Column 2 (B2)

        // Add a validation to the collection for the defined area
        int validationIndex = mainSheet.Validations.Add(validationArea);
        Validation validation = mainSheet.Validations[validationIndex];

        // Configure the validation to use a list sourced from the hidden sheet
        validation.Type = ValidationType.List;
        // Reference the hidden range using A1 notation with the sheet name
        validation.Formula1 = $"=HiddenData!$A$1:$A${allowedValues.Length}";
        validation.InCellDropDown = true; // Show the drop‑down arrow in the cell

        // Save the workbook
        workbook.Save("DataValidationFromHiddenSheet.xlsx");
    }
}
