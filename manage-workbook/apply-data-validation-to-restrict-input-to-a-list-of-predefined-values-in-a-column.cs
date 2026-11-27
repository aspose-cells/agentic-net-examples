// Title: Add a list‑type data validation dropdown to an Excel column using Aspose.Cells for .NET (C#)
// Common Searches: asp.net add list data validation to Excel column with Aspose.Cells C# | C# example for creating an in‑cell dropdown list (A2:A100) using Aspose.Cells | how to restrict column input to predefined values in an Excel workbook with Aspose.Cells .NET | Aspose.Cells set Validation.Type to List and specify Formula1 in C# code | generate Excel file with dropdown list validation using Aspose.Cells for .NET
// Tags: list‑type validation Aspose.Cells C# | in‑cell dropdown list .NET | Validation.Type List Aspose.Cells | Formula1 predefined options Aspose.Cells | export workbook with validation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, defines the cell range A2:A100, adds a list‑type validation with the options "Option1,Option2,Option3", enables an in‑cell dropdown, and saves the file as DataValidationExample.xlsx.
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

            // Define the range for the column (e.g., A2:A100)
            int firstRow = 1;   // Row 2 (zero‑based index)
            int lastRow = 99;   // Row 100
            int columnIndex = 0; // Column A

            // Create the cell area for the validation (startRow, startColumn, endRow, endColumn)
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = columnIndex,
                EndRow = lastRow,
                EndColumn = columnIndex
            };

            // Add a list validation to the worksheet for the specified area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;          // Restrict to a list
            validation.Operator = OperatorType.None;
            validation.InCellDropDown = true;               // Show drop‑down arrow

            // Define the allowed values (comma‑separated, enclosed in double quotes)
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Define output file path
            string outputPath = "DataValidationExample.xlsx";

            // Ensure the output directory exists
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
