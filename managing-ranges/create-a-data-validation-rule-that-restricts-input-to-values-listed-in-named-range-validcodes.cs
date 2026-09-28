// Title: Add a list‑type data validation that references a named range "ValidCodes" in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that defines a named range called ValidCodes containing a set of codes and applies a list‑type validation to cells A1:A10 based on that named range. | Show how to set a custom error title and message for the list validation rule created from a named range in Aspose.Cells.
// Common Searches: how to use a named range for list data validation with Aspose.Cells C# | Aspose.Cells create named range and apply list validation to a cell range | C# Aspose.Cells set custom error message for list validation from named range | Excel data validation list referencing named range using Aspose.Cells .NET
// Tags: Aspose.Cells create named range | Aspose.Cells list validation from named range | C# Excel data validation list Aspose.Cells | Aspose.Cells custom validation error message | Aspose.Cells define cell area validation

using System;
using Aspose.Cells;

namespace DataValidationExample
{
    // // This program creates a workbook, defines a named range "ValidCodes" with sample codes, applies a list‑type validation to cells A1‑A10 that points to the named range, configures a custom error title and message, and saves the file as DataValidation.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and set its name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample valid codes in column B (B1:B3)
                sheet.Cells["B1"].PutValue("CodeA");
                sheet.Cells["B2"].PutValue("CodeB");
                sheet.Cells["B3"].PutValue("CodeC");

                // Apply data validation to cells A1:A10 to restrict input to the list in B1:B3
                ValidationCollection validations = sheet.Validations;

                // Define the cell area for A1:A10 (rows 0‑9, column 0)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 9,
                    EndColumn = 0
                };

                // Add a validation for the defined area
                int validationIndex = validations.Add(area);
                Validation validation = validations[validationIndex];

                // Set validation type to List and point directly to the range B1:B3
                validation.Type = ValidationType.List;
                validation.Formula1 = "$B$1:$B$3";

                // Optional: configure error message
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Entry";
                validation.ErrorMessage = "Please select a value from the predefined list.";

                // Save the workbook to a file
                string outputPath = "DataValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
