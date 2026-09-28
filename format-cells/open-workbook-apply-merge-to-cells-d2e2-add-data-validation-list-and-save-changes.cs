// Title: How to merge cells D2:E2 and add a dropdown list validation to A1:A10 in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write a C# program that merges the range D2:E2 on the first worksheet, creates a list‑type data validation for cells A1 through A10 with custom error titles and messages, and then saves the workbook. | Modify the example so it works on a worksheet named "Sheet2" and changes the dropdown options to "Red,Green,Blue" while keeping the same merged cells and validation settings. | Generate code that loads an existing workbook, merges any specified cell range, adds a dropdown validation with a given comma‑separated list, and writes the updated file to a new location.
// Common Searches: Aspose.Cells C# merge specific cells and set a dropdown validation list | programmatically add list validation to A1:A10 after merging cells with Aspose.Cells | C# example for creating a merged header and attaching a data validation list in Excel | save changes to an Excel workbook after applying cell merging and validation using Aspose.Cells
// Tags: cell range merging Aspose.Cells | list validation creation Aspose.Cells | validation error handling Aspose.Cells | workbook saving Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads (or creates) an Excel workbook, merges cells D2:E2 on the first worksheet, adds a list‑type data validation dropdown to A1:A10 with custom error titles and messages, and saves the modified file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells D2:E2 (row index 1, column indices 3 to 4)
            sheet.Cells.Merge(1, 3, 1, 4);

            // Define the range for the validation (A1:A10)
            CellArea validationArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a data validation list to the specified range
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.Operator = OperatorType.None;
            // Define the list values (comma‑separated, enclosed in quotes)
            validation.Formula1 = "\"Option1,Option2,Option3\"";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please select a value from the list.";

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
