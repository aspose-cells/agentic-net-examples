// Title: Create a list‑based data validation drop‑down for cells A2:A20 in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that adds a List validation to range A2:A20, includes an input prompt, and saves the workbook. | Write a method that builds the quoted, comma‑separated string required for a list validation and applies it to a column using Aspose.Cells. | Show how to configure an error alert for a list‑type data validation in an Aspose.Cells worksheet and ensure the output folder exists before saving.
// Common Searches: Aspose.Cells C# how to add a drop‑down list validation to a specific column range | C# create list data validation for Excel cells A2 to A20 with Aspose.Cells | Set input message and error alert for Excel data validation using Aspose.Cells .NET | Save Excel workbook after applying list validation with Aspose.Cells in C#
// Tags: Aspose.Cells list validation C# | Excel drop-down validation Aspose.Cells | validation input message Aspose.Cells | error alert list validation .NET | save workbook after validation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, defines the range A2:A20 on the first worksheet, adds a List‑type validation containing "Apple, Banana, Cherry", sets an input message and an error alert, ensures the output directory exists, and saves the file as DataValidationExample.xlsx.
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

            // Define the range for data validation (A2:A20)
            // CellArea uses zero‑based indexes: row 1 = A2, row 19 = A20, column 0 = A
            CellArea validationArea = new CellArea
            {
                StartRow = 1,
                StartColumn = 0,
                EndRow = 19,
                EndColumn = 0
            };

            // List of allowed values
            string[] allowedValues = { "Apple", "Banana", "Cherry" };
            // Aspose.Cells expects the list as a quoted, comma‑separated string
            string list = $"\"{string.Join(",", allowedValues)}\"";

            // Add a new validation object for the specified area
            ValidationCollection validations = sheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];

            // Set validation type and allowed values
            validation.Type = ValidationType.List;
            validation.Formula1 = list;

            // Optional: display input message and error alert
            validation.InputMessage = "Select a fruit from the list.";
            validation.ErrorMessage = "Invalid entry. Choose a value from the list.";
            validation.ShowError = true; // show error alert when validation fails

            // Save the workbook
            string outputPath = "DataValidationExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
