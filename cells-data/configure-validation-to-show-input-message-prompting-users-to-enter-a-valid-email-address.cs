// Title: Add email address validation with input prompt and error alert to cell B2 in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to add email address validation with a prompt message in Aspose.Cells C# | Aspose.Cells custom validation formula for checking email format in .NET | set data validation input message and error alert for a single cell using Aspose.Cells | C# create Excel file with email validation and custom error title using Aspose.Cells | apply custom validation rule to a cell range in Aspose.Cells .NET example
// Tags: email validation formula Aspose.Cells | input message for data validation .NET | error alert configuration Excel C# | cell-specific validation range Aspose.Cells | Excel workbook creation with validation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, defines cell B2, adds a custom validation rule that verifies the presence of '@' and a following '.' using an Excel formula, configures an input prompt and an error alert with a custom title and message, and saves the file as EmailValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range where email validation is required (e.g., B2)
            CellArea area = new CellArea
            {
                StartRow = 1,   // Row index is zero‑based (1 = second row)
                EndRow = 1,
                StartColumn = 1, // Column index is zero‑based (1 = column B)
                EndColumn = 1
            };

            // Add a custom validation rule to the defined range
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.Custom;

            // Formula checks that the cell contains an "@" character and a "." after it
            validation.Formula1 = "AND(ISNUMBER(FIND(\"@\",B2)), ISNUMBER(FIND(\".\",B2, FIND(\"@\",B2))))";

            // Optional: configure an error alert for invalid entries
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Email";
            validation.ErrorMessage = "The value entered is not a valid email address.";

            // Determine output file path
            string outputPath = "EmailValidation.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
