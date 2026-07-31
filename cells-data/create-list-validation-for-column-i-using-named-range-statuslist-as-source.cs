// Title: C# – Add List Validation to Column I Using Named Range "StatusList" with Aspose.Cells
// Description: Creates a workbook, defines a named range called StatusList (A1:A3), and applies a List‑type data validation to column I (rows 1‑101). The validation references the named range, shows an in‑cell dropdown, and displays a custom input title and message.
// Keywords: Aspose.Cells C# list validation | named range validation .NET | dropdown list column I Aspose.Cells | ValidationType.List named range | StatusList named range | Aspose.Cells data validation example | C# Excel dropdown from named range
// Common Searches: Aspose.Cells add dropdown list to column using named range | C# Aspose.Cells list validation column I | How to reference a named range in Aspose.Cells validation | Aspose.Cells set data validation for multiple rows | Create named range and apply list validation Aspose.Cells
// Developer Intent: Apply a List‑type validation to column I that pulls allowed values from the existing named range "StatusList".
// Use Cases: Standardize status entry in task‑tracking sheets without hard‑coding values. | Reuse a single named range across several worksheets to keep validation consistent. | Generate reports programmatically where users must select a status from a predefined list.
// AI Prompts: Generate C# Aspose.Cells code to add list validation to column J using an existing named range "CategoryList" with a custom input title and message. | Show how to expand a validation range so the dropdown applies to the entire column dynamically in Aspose.Cells. | Explain how to reference a pre‑defined named range in a validation formula when the workbook already contains that range.

using System;
using Aspose.Cells;

// Creates a workbook, defines a named range called StatusList (A1:A3), and applies a List‑type data validation to column I (rows 1‑101). The validation references the named range, shows an in‑cell dropdown, and displays a custom input title and message.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Create a sample list that will be used as the source for the
            // named range "StatusList". In a real scenario the named range
            // might already exist, so this part can be omitted.
            // ------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Open");
            worksheet.Cells["A2"].PutValue("Closed");
            worksheet.Cells["A3"].PutValue("Pending");

            // Define the named range "StatusList" that refers to A1:A3
            int nameIndex = workbook.Worksheets.Names.Add("StatusList");
            Name statusName = workbook.Worksheets.Names[nameIndex];
            statusName.RefersTo = "=$A$1:$A$3";

            // ------------------------------------------------------------
            // Add list validation to column I (zero‑based index 8)
            // The validation will apply to rows 0 through 100 (adjust as needed)
            // ------------------------------------------------------------
            CellArea validationArea = new CellArea
            {
                StartRow = 0,      // first row (0‑based)
                StartColumn = 8,   // column I
                EndRow = 100,      // last row
                EndColumn = 8      // column I
            };

            // Add the validation to the worksheet's validation collection
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation as a List that uses the named range
            validation.Type = ValidationType.List;
            validation.Formula1 = "StatusList";   // reference to the named range
            validation.InCellDropDown = true;    // show dropdown arrow
            validation.ShowInput = true;         // display input message
            validation.InputTitle = "Status";
            validation.InputMessage = "Select a status from the list.";

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("ListValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
