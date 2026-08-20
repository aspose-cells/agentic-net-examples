// Title: C# – Add List Data Validation from a Named Range with Aspose.Cells
// Description: Demonstrates how to create a workbook, define a named range (AllowedRange) covering cells A1:A3, and apply a list‑type data validation to cell B1 that references the named range, showing an in‑cell drop‑down and custom error message, then saves the file.
// Keywords: Aspose.Cells | C# | list data validation | named range | Excel dropdown | data validation formula | Aspose.Cells .NET | restrict cell input | validation with named range | Excel automation
// Common Searches: Aspose.Cells list validation named range C# | How to restrict Excel cell input using Aspose.Cells | Create named range and apply data validation in C# | Aspose.Cells drop‑down list from named range | C# Excel data validation with Aspose.Cells
// Developer Intent: Implement a data‑validation rule that limits cell entry to the values defined in a specific named range.
// Use Cases: Ensure users select only predefined options in a form cell | Reuse the same named‑range validation across multiple columns | Prevent accidental formula entry by allowing only list values | Generate templates with locked dropdown choices for data collection
// AI Prompts: Write C# code using Aspose.Cells to create a named range and attach a list validation to a cell range. | Show how to change the validation to reference a different named range passed as a variable. | Explain how to export the workbook with validation to a MemoryStream instead of a file. | Provide code to apply the same named‑range validation to an entire column. | Describe how to customize the error alert title and message for the validation.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, define a named range (AllowedRange) covering cells A1:A3, and apply a list‑type data validation to cell B1 that references the named range, showing an in‑cell drop‑down and custom error message, then saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data that will be used in the named range
        sheet.Cells["A1"].PutValue("Option1");
        sheet.Cells["A2"].PutValue("Option2");
        sheet.Cells["A3"].PutValue("Option3");

        // -----------------------------------------------------------------
        // 1. Define a named range called "AllowedRange" that refers to A1:A3
        // -----------------------------------------------------------------
        int nameIndex = workbook.Worksheets.Names.Add("AllowedRange");
        Name allowedRange = workbook.Worksheets.Names[nameIndex];
        // The RefersTo string must start with '=' and include the sheet name
        allowedRange.RefersTo = "=Sheet1!$A$1:$A$3";

        // ---------------------------------------------------------------
        // 2. Add a data validation that restricts entry to the named range
        // ---------------------------------------------------------------
        // Define the cell (B1) where the validation will be applied
        CellArea validationArea = CellArea.CreateCellArea(0, 1, 0, 1); // Row 0, Column 1 => B1

        // Add a new validation to the worksheet for the specified area
        int validationIndex = sheet.Validations.Add(validationArea);
        Validation validation = sheet.Validations[validationIndex];

        // Use a List type validation and point Formula1 to the named range
        validation.Type = ValidationType.List;
        validation.Formula1 = "=AllowedRange";   // Must start with '='
        validation.InCellDropDown = true;       // Show drop‑down list
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Entry";
        validation.ErrorMessage = "Please select a value from the predefined list.";

        // ---------------------------------------------------------------
        // 3. Save the workbook
        // ---------------------------------------------------------------
        workbook.Save("ValidationWithNamedRange.xlsx");
    }
}
