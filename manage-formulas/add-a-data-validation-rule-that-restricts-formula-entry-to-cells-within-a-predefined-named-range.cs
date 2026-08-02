// Title: Add List Data Validation with a Named Range using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate cells A1‑A5, define a workbook‑level named range "AllowedValues", and apply a List‑type data validation to cell B1 that references the named range. Includes optional drop‑down, custom error title, and message, then saves the file as ValidationWithNamedRange.xlsx.
// Keywords: Aspose.Cells C# data validation list | named range validation .NET | list validation Aspose.Cells | restrict cell input named range | Aspose.Cells ValidationType.List example | C# Excel drop‑down validation
// Common Searches: Aspose.Cells list validation using named range | C# restrict cell entry to predefined list Aspose.Cells | How to add data validation list from named range in .NET | Aspose.Cells create named range and apply validation | Excel drop‑down validation with Aspose.Cells C#
// Developer Intent: Implement a data‑validation rule that limits cell input to the values defined in a named range.
// Use Cases: Provide users with a drop‑down of allowed options sourced from a central list. | Prevent invalid formulas or free‑text entries by enforcing a predefined list. | Reuse the same validation across multiple worksheets by referencing a workbook‑level named range.
// AI Prompts: Generate C# code with Aspose.Cells that defines a named range and adds a List validation referencing it, including custom error title and message. | Show how to extend the named‑range list validation to a range of cells (e.g., B2:B20) with in‑cell drop‑down enabled.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, populate cells A1‑A5, define a workbook‑level named range "AllowedValues", and apply a List‑type data validation to cell B1 that references the named range. Includes optional drop‑down, custom error title, and message, then saves the file as ValidationWithNamedRange.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // 1. Populate some sample data that will be used as the allowed list
        // ------------------------------------------------------------
        sheet.Cells["A1"].PutValue("Alpha");
        sheet.Cells["A2"].PutValue("Beta");
        sheet.Cells["A3"].PutValue("Gamma");
        sheet.Cells["A4"].PutValue("Delta");
        sheet.Cells["A5"].PutValue("Epsilon");

        // ------------------------------------------------------------
        // 2. Define a named range that refers to the cells A1:A5
        // ------------------------------------------------------------
        int nameIdx = workbook.Worksheets.Names.Add("AllowedValues");
        Name allowedRange = workbook.Worksheets.Names[nameIdx];
        // The RefersTo string must start with '=' and use A1 notation
        allowedRange.RefersTo = "=Sheet1!$A$1:$A$5";

        // ------------------------------------------------------------
        // 3. Add a data validation to cell B1 that restricts entry to the named range
        // ------------------------------------------------------------
        // Define the area where the validation will be applied (B1)
        CellArea validationArea = CellArea.CreateCellArea(0, 1, 0, 1); // Row 0, Column 1 => B1

        // Add the validation to the worksheet's ValidationCollection
        int validationIdx = sheet.Validations.Add(validationArea);
        Validation validation = sheet.Validations[validationIdx];

        // Use a List type validation and point Formula1 to the named range
        validation.Type = ValidationType.List;
        validation.Formula1 = "=AllowedValues";

        // Optional UI settings
        validation.InCellDropDown = true;          // Show drop‑down arrow
        validation.ShowError = true;               // Show error dialog on invalid entry
        validation.ErrorTitle = "Invalid Entry";
        validation.ErrorMessage = "Please select a value from the predefined list.";

        // ------------------------------------------------------------
        // 4. Save the workbook
        // ------------------------------------------------------------
        workbook.Save("ValidationWithNamedRange.xlsx");
    }
}
