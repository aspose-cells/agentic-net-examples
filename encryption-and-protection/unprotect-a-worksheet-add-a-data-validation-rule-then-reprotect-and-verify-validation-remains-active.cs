// Title: C# Example: Unprotect Worksheet, Add List‑Dropdown Validation, Re‑protect with Aspose.Cells
// Description: This Aspose.Cells for .NET sample creates a workbook, protects the first worksheet with a password, unprotects it, adds a list‑type data validation to cell A1, confirms the rule exists, re‑protects the sheet using the same password, verifies the validation remains, and saves the file as WorksheetValidationProtected.xlsx.
// Keywords: Aspose.Cells | C# | .NET | worksheet protection | unprotect worksheet | list data validation | dropdown validation | Excel validation persistence | password protection | sample code | GitHub example | data validation after protect
// Common Searches: how to add a dropdown list to a protected sheet using Aspose.Cells C# | preserve data validation after re‑protecting an Excel worksheet | unprotect and protect worksheet programmatically with password Aspose.Cells | C# code to verify validation exists after worksheet protection | Aspose.Cells example for worksheet protection and validation
// Developer Intent: Add a list‑type validation to a cell on a protected worksheet and ensure the rule stays active after re‑protecting, using Aspose.Cells in C#.
// Use Cases: Generate a locked template that still offers users a predefined dropdown for data entry. | Automate workbook preparation where validation rules must be added or updated before distribution while keeping the sheet secured. | Create reports that require certain cells to be immutable but need selectable options via Excel’s data‑validation dropdown.
// AI Prompts: Provide C# code to add a date‑range validation to a cell after unprotecting a worksheet and keep it after re‑protecting with Aspose.Cells. | Show how to modify an existing data‑validation rule on a protected sheet and re‑apply protection without losing other validations. | Explain how to customize error and input messages for dropdown validations on a password‑protected worksheet saved with Aspose.Cells.

using System;
using Aspose.Cells;

// This Aspose.Cells for .NET sample creates a workbook, protects the first worksheet with a password, unprotects it, adds a list‑type data validation to cell A1, confirms the rule exists, re‑protects the sheet using the same password, verifies the validation remains, and saves the file as WorksheetValidationProtected.xlsx.
class WorksheetValidationProtectionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        string password = "Secret123";

        // Protect the worksheet initially
        worksheet.Protect(ProtectionType.All, password, null);
        Console.WriteLine("Worksheet initially protected: " + worksheet.IsProtected);

        // Unprotect the worksheet using the password
        worksheet.Unprotect(password);
        Console.WriteLine("Worksheet after unprotect: " + !worksheet.IsProtected);

        // Add a list‑type data validation to cell A1
        CellArea validationArea = CellArea.CreateCellArea(0, 0, 0, 0); // A1
        int validationIndex = worksheet.Validations.Add(validationArea);
        Validation validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.Formula1 = "Option1,Option2,Option3";
        validation.InCellDropDown = true;
        validation.ShowInput = true;
        validation.InputMessage = "Select an option";
        validation.ErrorMessage = "Invalid selection";

        // Verify that the validation exists before re‑protecting
        Validation beforeProtect = worksheet.Validations.GetValidationInCell(0, 0);
        Console.WriteLine("Validation type before protect: " + (beforeProtect != null ? beforeProtect.Type.ToString() : "None"));

        // Re‑protect the worksheet with the same password
        worksheet.Protect(ProtectionType.All, password, null);
        Console.WriteLine("Worksheet re‑protected: " + worksheet.IsProtected);

        // Verify that the validation still exists after protection
        Validation afterProtect = worksheet.Validations.GetValidationInCell(0, 0);
        Console.WriteLine("Validation type after protect: " + (afterProtect != null ? afterProtect.Type.ToString() : "None"));

        // Save the workbook
        workbook.Save("WorksheetValidationProtected.xlsx");
    }
}
