// Title: C# – Convert CSV to XLSX and Apply Whole‑Number Data Validation with Aspose.Cells
// Description: This example demonstrates how to use Aspose.Cells for .NET to convert a CSV file to an XLSX workbook, add a data‑validation rule that limits values in column B (rows 2‑100) to whole numbers between 1 and 100, and save the result with OoxmlSaveOptions.
// Keywords: Aspose.Cells | C# | CSV to XLSX conversion | data validation | whole number range | ValidationType.WholeNumber | ConversionUtility | OoxmlSaveOptions | Excel column validation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add numeric validation after CSV conversion | C# restrict column values in Excel with Aspose.Cells | How to set data validation for a range in Aspose.Cells | Convert CSV to XLSX and enforce whole number limits | Aspose.Cells ValidationCollection example
// Developer Intent: Create an XLSX file from a CSV and enforce a numeric range constraint on a specific column using Aspose.Cells for .NET.
// Use Cases: Ensuring imported quantity fields stay within allowed limits | Validating ID numbers from a CSV before further processing | Protecting template worksheets from out‑of‑range entries | Automating data quality checks in financial reports generated from CSV sources | Applying consistent input rules across multiple generated spreadsheets
// AI Prompts: Write C# code with Aspose.Cells that converts a CSV to XLSX and adds a whole‑number validation to column C rows 5‑200. | Explain how to change the validation to a drop‑down list of text values in Aspose.Cells. | Show a loop that applies the same numeric validation to columns D through G using Aspose.Cells. | Provide steps to customize the input and error messages for data validation in an Aspose.Cells workbook. | Generate a PowerShell script that calls Aspose.Cells to perform the same conversion and validation.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// This example demonstrates how to use Aspose.Cells for .NET to convert a CSV file to an XLSX workbook, add a data‑validation rule that limits values in column B (rows 2‑100) to whole numbers between 1 and 100, and save the result with OoxmlSaveOptions.
class Program
{
    static void Main()
    {
        // Paths for source CSV and target XLSX files
        string csvPath = "input.csv";
        string xlsxPath = "output.xlsx";

        // 1. Convert CSV to XLSX using the provided ConversionUtility.Convert method
        ConversionUtility.Convert(csvPath, xlsxPath);

        // 2. Load the converted workbook (standard load)
        Workbook workbook = new Workbook(xlsxPath);

        // 3. Access the first worksheet where validation will be added
        Worksheet sheet = workbook.Worksheets[0];

        // 4. Define the cell area for the validation (e.g., column B, rows 2‑100)
        CellArea validationArea = new CellArea
        {
            StartRow = 1,      // Row 2 (0‑based index)
            EndRow = 99,       // Row 100
            StartColumn = 1,   // Column B (0‑based index)
            EndColumn = 1
        };

        // 5. Add a new validation rule to the worksheet
        ValidationCollection validations = sheet.Validations;
        int validationIndex = validations.Add(validationArea);
        Validation validation = validations[validationIndex];

        // 6. Configure the validation (whole numbers between 1 and 100)
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";
        validation.Formula2 = "100";
        validation.InputMessage = "Please enter a whole number between 1 and 100.";
        validation.ErrorMessage = "Invalid entry. Value must be between 1 and 100.";
        validation.ShowInput = true;
        validation.ShowError = true;

        // 7. Save the workbook using OoxmlSaveOptions (provided SaveOptions class)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        workbook.Save(xlsxPath, saveOptions);
    }
}
