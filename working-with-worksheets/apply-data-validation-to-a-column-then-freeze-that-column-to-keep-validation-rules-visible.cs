// Title: Create an Excel file with list‑type data validation in column A and freeze that column using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a list‑based data validation with an in‑cell dropdown to the range A1:A100 and then freezes column A using Aspose.Cells. | Generate a .NET program that configures a validation list, sets a custom error title and message, enables the dropdown arrow, and applies FreezePanes to keep the validation column visible while scrolling.
// Common Searches: how to add a dropdown list validation to column A with Aspose.Cells C# | Aspose.Cells freeze first column after applying data validation | C# example for list validation and FreezePanes in the same worksheet | apply data validation list and freeze panes together using Aspose.Cells for .NET | Aspose.Cells create workbook with validation dropdown and frozen column
// Tags: Aspose.Cells list‑type validation dropdown C# | Aspose.Cells FreezePanes column A .NET | Aspose.Cells custom validation error message | Aspose.Cells apply validation to range A1:A100 | Aspose.Cells workbook creation with validation and frozen column

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new workbook, adds a list‑type data validation with an in‑cell dropdown to cells A1:A100, configures custom error title and message, freezes column A using FreezePanes, and saves the file as DataValidationAndFreeze.xlsx.
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

            // Define the range for data validation (e.g., column A, rows 1-100)
            int firstRow = 0;          // Row index is zero‑based (row 1)
            int lastRow = 99;          // Row 100
            int columnIndex = 0;       // Column A

            // Build the Excel address for the validation range (e.g., "A1:A100")
            string startCell = CellsHelper.CellIndexToName(firstRow, columnIndex);
            string endCell = CellsHelper.CellIndexToName(lastRow, columnIndex);

            // Convert the address strings to a CellArea object required by Add()
            CellArea validationArea = CellArea.CreateCellArea(startCell, endCell);

            // Add a list‑type validation to the specified range
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;                     // List validation
            validation.Operator = OperatorType.Equal;                  // Required for list type
            validation.Formula1 = "\"Option1,Option2,Option3\"";       // Comma‑separated list
            validation.ShowError = true;                               // Show error dialog
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please select a value from the list.";
            validation.InCellDropDown = true;                          // Show dropdown arrow (correct property)

            // Freeze the first column (column A) so validation rules stay visible while scrolling
            // Parameters: row, column, rowOffset, columnOffset
            // Setting column = 1 freezes columns to the left of column B (i.e., column A)
            sheet.FreezePanes(0, 1, 0, 0);

            // Save the workbook to a file
            string outputPath = "DataValidationAndFreeze.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
