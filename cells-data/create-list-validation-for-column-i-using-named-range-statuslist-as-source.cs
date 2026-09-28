// Title: Add a drop‑down list validation to column I using a named range (StatusList) with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a hidden worksheet, defines a named range called StatusList with status values, and applies it as a list‑type data validation to column I of the first sheet using Aspose.Cells. | Generate an Aspose.Cells example that populates status options on a separate sheet, creates a named range for those cells, and sets a drop‑down validation for cells I1:I1000 in the main worksheet.
// Common Searches: aspocells c# create named range for data validation column i | how to add drop down list to excel column i using aspocells and a hidden sheet | c# aspocells list validation referencing a named range statuslist | excel data validation list from another worksheet aspocells .net example | aspocells set validation range for column i rows 1 to 1000
// Tags: Aspose.Cells named range list validation | C# hidden worksheet validation source | Aspose.Cells set drop‑down validation Excel | Excel column I data validation using Aspose.Cells | Aspose.Cells define StatusList range

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, adds a hidden sheet named "Lists" with status values in Z1:Z3, defines a CellArea covering column I rows 1‑1000, adds a list‑type validation that references the range on the hidden sheet (which can be defined as the named range "StatusList"), sets input and error messages, and saves the workbook as ListValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet where validation will be applied
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Create a hidden sheet to store the list values
            // -------------------------------------------------
            int listSheetIndex = workbook.Worksheets.Add(); // add a new sheet
            Worksheet listSheet = workbook.Worksheets[listSheetIndex];
            listSheet.Name = "Lists";

            // Populate the list (example values)
            listSheet.Cells["Z1"].PutValue("Open");
            listSheet.Cells["Z2"].PutValue("In Progress");
            listSheet.Cells["Z3"].PutValue("Closed");

            // -------------------------------------------------
            // Apply list validation to column I (zero‑based index 8)
            // -------------------------------------------------
            int columnI = 8;               // Column I (zero‑based)
            int startRow = 0;              // Row 1 (zero‑based)
            int endRow = 999;              // Row 1000 (adjust as needed)

            // Define the area for validation
            CellArea area = new CellArea
            {
                StartRow = startRow,
                StartColumn = columnI,
                EndRow = endRow,
                EndColumn = columnI
            };

            // Add validation for the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            validation.Type = ValidationType.List;      // List validation
            validation.Operator = OperatorType.None;    // Not used for list type

            // Reference the list range directly (no named range needed)
            validation.Formula1 = $"'{listSheet.Name}'!$Z$1:$Z$3";

            // Optional: user messages
            validation.InputMessage = "Select a status from the list.";
            validation.ErrorMessage = "Invalid entry. Choose a value from the predefined list.";

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "ListValidation.xlsx";

            // Ensure the output directory exists
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
