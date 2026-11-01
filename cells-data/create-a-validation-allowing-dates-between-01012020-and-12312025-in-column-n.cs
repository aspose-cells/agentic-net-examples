// Title: Add date validation to column N for dates between 01/01/2020 and 12/31/2025 with Aspose.Cells in C#
// AI Prompts: Write C# code using Aspose.Cells that adds a date validation to column N, allowing only dates from 01/01/2020 to 12/31/2025. | Show how to configure custom input and error messages for a date‑range validation applied to an entire worksheet column with Aspose.Cells. | Create a reusable C# method that applies a Between‑date validation to any specified column in a workbook using Aspose.Cells.
// Common Searches: asp.net add date validation to column N using Aspose.Cells | c# restrict Excel column N dates between 2020 and 2025 with Aspose.Cells | how to set custom input and error messages for date validation in Aspose.Cells | apply whole‑column date range validation in .NET Excel library Aspose.Cells
// Tags: Aspose.Cells date validation between specific dates | C# apply validation to entire column N | Excel column N date range restriction using Aspose.Cells | Aspose.Cells custom input and error messages for validation | OperatorType.Between usage in Aspose.Cells validations

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new workbook, adds a date validation to the full column N that permits only dates from 01/01/2020 to 12/31/2025, sets optional input and error messages, and saves the file as DateValidation.xlsx.
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

            // Define the range for column N (N1:N1048576)
            CellArea range = CellArea.CreateCellArea("N1", "N1048576");

            // Add a date validation to the specified range.
            // In some Aspose.Cells versions Add returns the index of the new validation.
            int validationIndex = sheet.Validations.Add(range);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to Date
            validation.Type = ValidationType.Date;

            // Require the date to be between the two specified dates
            validation.Operator = OperatorType.Between;

            // Set the lower and upper bounds (Excel date format)
            validation.Formula1 = "01/01/2020";
            validation.Formula2 = "12/31/2025";

            // Optional: Show input message and error alert
            validation.InputMessage = "Enter a date between 01/01/2020 and 12/31/2025.";
            validation.ErrorMessage = "Invalid date. Please enter a date within the allowed range.";
            // Uncomment if your Aspose.Cells version supports these properties
            // validation.ShowInputMessage = true;
            // validation.ShowError = true;

            // Determine output path and ensure its directory exists
            string outputPath = "DateValidation.xlsx";
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
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
