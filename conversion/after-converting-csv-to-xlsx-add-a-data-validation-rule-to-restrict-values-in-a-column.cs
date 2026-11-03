// Title: Convert a CSV file to XLSX and apply a list‑type data validation to column B using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add dropdown list validation to column B after converting CSV to Excel | How to set ValidationType.List for a range of cells in a workbook created from CSV with Aspose.Cells | C# code to clear existing validations and create a list validation in B2:B100 using Aspose.Cells | Convert CSV to XLSX and restrict values with a data validation list in Aspose.Cells .NET example
// Tags: CSV to XLSX conversion with Aspose.Cells | list data validation in Aspose.Cells C# | apply ValidationType.List to cell range using Aspose.Cells | clear worksheet validations Aspose.Cells | configure validation error message Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a CSV file (or creates a new workbook if missing), clears any existing validations, adds a list‑type validation with predefined options to cells B2:B100, sets an error title and message, and saves the result as an XLSX workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.csv";
            string outputPath = "output.xlsx";

            // Load the CSV file if it exists; otherwise create a new workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create an empty workbook with a single worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Get the first worksheet (the CSV data is loaded here)
            Worksheet sheet = workbook.Worksheets[0];

            // Access the collection of data validations for the worksheet
            ValidationCollection validations = sheet.Validations;

            // Optional: clear existing validations
            validations.Clear();

            // Define the cell range for the validation (B2:B100)
            CellArea area = new CellArea
            {
                StartRow = 1,      // Row 2 (0‑based index)
                StartColumn = 1,   // Column B (0‑based index)
                EndRow = 99,       // Row 100
                EndColumn = 1      // Column B
            };

            // Add a new validation rule for the specified area
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];

            // Set the validation type to a list (drop‑down) and define allowed values
            validation.Type = ValidationType.List;
            validation.Operator = OperatorType.Equal;
            // The list of allowed values (comma‑separated) must be enclosed in double quotes
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Configure the error message that appears when an invalid value is entered
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid entry";
            validation.ErrorMessage = "Please select a value from the list.";

            // Save the workbook as an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
