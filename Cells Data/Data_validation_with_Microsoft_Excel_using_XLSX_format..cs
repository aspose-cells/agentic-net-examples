using System;
using System.IO;
using Aspose.Cells;

class DataValidationDemo
{
    static void Main()
    {
        // Create a new workbook (default XLSX format)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a whole‑number validation that requires values between 10 and 100
        int index = worksheet.Validations.Add();
        Validation validation = worksheet.Validations[index];
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "10";
        validation.Formula2 = "100";

        // Apply the validation to the range A1:A10
        CellArea area = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based index)
            StartColumn = 0,   // Column A
            EndRow = 9,        // Row 10
            EndColumn = 0      // Column A
        };
        validation.AddArea(area);

        // Ensure the output file is not locked by a previous run
        string outputPath = "DataValidationDemo.xlsx";
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        // Save the workbook as an XLSX file
        workbook.Save(outputPath);
    }
}