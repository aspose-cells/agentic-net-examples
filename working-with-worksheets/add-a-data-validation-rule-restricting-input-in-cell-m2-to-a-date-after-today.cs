// Title: Add a future‑date validation rule to cell M2 in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that applies a date validation to cell M2, allowing only dates later than the current day and displaying a custom error message. | Create an Aspose.Cells workbook where the validation for M2 uses the GreaterThan operator with today's date as the lower bound. | Write a C# snippet that adds a ValidationType.Date rule to M2, sets OperatorType.GreaterThan, assigns Formula1 to DateTime.Today, and saves the file.
// Common Searches: Aspose.Cells C# how to set date validation for a specific cell to only accept future dates | C# add data validation to Excel cell M2 that restricts input to dates after today using Aspose.Cells | example of using OperatorType.GreaterThan with ValidationType.Date in Aspose.Cells | Aspose.Cells create custom error message for a date validation rule
// Tags: Aspose.Cells date validation future dates | C# Aspose.Cells ValidationType.Date | OperatorType.GreaterThan Excel validation | custom error message Aspose.Cells validation | CellArea M2 Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example creates a new workbook, defines a CellArea for M2, adds a date‑type validation that only permits dates greater than today, configures a custom error title and message, ensures the output directory exists, and saves the workbook as Output.xlsx.
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

            // Add a data validation rule for cell M2
            ValidationCollection validations = sheet.Validations;

            // Create CellArea for M2 (row index 1, column index 12)
            // Use overload that accepts start and end cell addresses
            CellArea area = CellArea.CreateCellArea("M2", "M2");
            int index = validations.Add(area);
            Validation validation = validations[index];

            // Set validation type to Date and restrict to dates greater than today
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.GreaterThan;
            validation.Formula1 = DateTime.Today.ToOADate().ToString(CultureInfo.InvariantCulture);

            // Optional: configure error message
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "Please enter a date after today.";

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the output directory exists (handle case when outputPath has no directory)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
