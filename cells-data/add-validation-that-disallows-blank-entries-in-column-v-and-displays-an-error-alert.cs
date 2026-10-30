// Title: How to add a custom data‑validation rule that blocks blank entries in column V and shows an error alert with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add data validation to prevent empty cells in column V | C# Aspose.Cells custom validation formula to disallow blanks and display an error message | Set IgnoreBlank false and show an error alert for a specific column using Aspose.Cells .NET | Create Excel workbook with column‑specific validation that blocks blank entries in Aspose.Cells | Aspose.Cells validation error title and message for column V in C#
// Tags: Aspose.Cells custom validation rule C# | disable blank entries Excel column using Aspose.Cells | configure validation error message Aspose.Cells .NET | apply validation to entire column V Aspose.Cells | data validation with LEN(TRIM) formula Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, defines the full column V (index 21) as a CellArea, adds a custom data‑validation rule, sets IgnoreBlank to false, uses the formula LEN(TRIM(V1))>0 to ensure cells are not empty or whitespace, configures an error alert with a title and message, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column V (zero‑based column index 21)
            CellArea columnV = new CellArea
            {
                StartRow = 0,
                EndRow = 1048575, // maximum rows in Excel
                StartColumn = 21,
                EndColumn = 21
            };

            // Add a data‑validation rule for the defined range
            int validationIndex = sheet.Validations.Add(columnV);
            Validation validation = sheet.Validations[validationIndex];

            // Disallow blank entries
            validation.IgnoreBlank = false;

            // Use a custom formula to ensure the cell is not empty or whitespace only
            validation.Type = ValidationType.Custom;
            validation.Formula1 = "LEN(TRIM(V1))>0";

            // Configure the error alert
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Entry";
            validation.ErrorMessage = "Blank entries are not allowed in column V.";

            // Determine output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
