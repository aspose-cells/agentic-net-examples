// Title: Create an Aspose.Cells workbook, add whole‑number validation to column A, and save the sheet as a JSON file in C#
// Common Searches: how to add whole number validation to a column using Aspose.Cells C# and export to JSON | Aspose.Cells C# save worksheet with data validation as JSON file | C# example for setting numeric range validation on column A and converting workbook to JSON | export Aspose.Cells workbook to JSON after applying validation rules | Aspose.Cells .NET create workbook, apply validation, save as JSON
// Tags: whole number validation Aspose.Cells C# | export worksheet to JSON Aspose.Cells | validation range A2:A100 Aspose.Cells | save workbook as JSON .NET | numeric range validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample program creates a new Workbook, renames the first worksheet, fills sample data, defines a validation area covering A2:A100, adds a whole‑number validation that restricts values to 1‑100, and then saves the workbook to a JSON file, creating the output directory if needed and handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["B3"].PutValue("Bob");

            // Define validation range (A2:A100)
            int firstRow = 2; // 1‑based index (excluding header)
            int lastRow = 100;
            int column = 0;   // column A (zero‑based)

            CellArea validationArea = new CellArea
            {
                StartRow = firstRow - 1,
                EndRow = lastRow - 1,
                StartColumn = column,
                EndColumn = column
            };

            // Add data validation to the worksheet
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];

            // Configure validation: whole numbers between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Optional: input and error messages (available in newer API versions)
            // If the current Aspose.Cells version does not support these properties,
            // they can be omitted safely.
            // validation.ShowInputMessage = true;
            // validation.InputTitle = "Enter ID";
            // validation.InputMessage = "Please enter a whole number between 1 and 100.";
            // validation.ShowError = true;
            // validation.ErrorTitle = "Invalid ID";
            // validation.ErrorMessage = "The ID must be a whole number between 1 and 100.";

            // Export the workbook to a JSON file
            string jsonPath = "DataSheet.json";

            // Ensure the directory exists (if a directory part is present)
            string jsonDir = Path.GetDirectoryName(Path.GetFullPath(jsonPath));
            if (!string.IsNullOrEmpty(jsonDir) && !Directory.Exists(jsonDir))
            {
                Directory.CreateDirectory(jsonDir);
            }

            workbook.Save(jsonPath, SaveFormat.Json);
            Console.WriteLine($"Workbook exported to JSON at: {jsonPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
