using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a CellArea that covers rows 2‑100 (zero‑based indices 1‑99) in column A (column index 0)
            CellArea area = CellArea.CreateCellArea(1, 0, 99, 0);

            // Add a validation for the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation as whole‑number type
            validation.Type = ValidationType.WholeNumber;

            // Aspose.Cells versions prior to 20.9 do not have OperatorType.GreaterThanOrEqual.
            // Use GreaterThan with a threshold of -1 to achieve ">= 0".
            validation.Operator = OperatorType.GreaterThan;
            validation.Formula1 = "-1";

            // Determine output path
            string outputPath = "WholeNumberValidation.xlsx";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}