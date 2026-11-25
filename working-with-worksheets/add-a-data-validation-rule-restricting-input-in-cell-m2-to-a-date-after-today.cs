// Title: Create a date‑validation rule in Aspose.Cells for .NET that allows only dates after today in cell M2
// Common Searches: Aspose.Cells C# add data validation for future dates in a specific cell | How to restrict Excel cell M2 to dates later than today using Aspose.Cells | Set validation formula TODAY() with GreaterThan operator in Aspose.Cells .NET | C# code to apply date validation to a single cell in a generated workbook | Aspose.Cells validation collection example for date after today
// Tags: Aspose.Cells date validation with TODAY() | C# Excel cell M2 validation rule | Aspose.Cells GreaterThan operator for dates | Data validation collection in Aspose.Cells .NET | Future date restriction in generated Excel workbook

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, defines the M2 cell area, adds a date‑type validation with the GreaterThan operator and a TODAY() formula, sets input and error messages, ensures the output directory exists, and saves the file as Output.xlsx.
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

            // Define the cell area for M2 (zero‑based indices: row 1, column 12)
            CellArea area = new CellArea
            {
                StartRow = 1,
                EndRow = 1,
                StartColumn = 12,
                EndColumn = 12
            };

            // Add a validation rule for the defined area
            ValidationCollection validations = sheet.Validations;
            int index = validations.Add(area);
            Validation validation = validations[index];

            // Set validation type to Date and require a date greater than today
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.GreaterThan;
            validation.Formula1 = "TODAY()";

            // Input prompt (shown automatically when the cell is selected)
            validation.InputTitle = "Enter Date";
            validation.InputMessage = "Please enter a date after today.";

            // Error message (shown when validation fails)
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "The date must be after today.";

            // Determine output path and ensure the directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
