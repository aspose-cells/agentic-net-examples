// Title: Create a new workbook, add whole‑number data validation to column B, and export the sheet as JSON using Aspose.Cells for .NET
// Common Searches: how to add whole number data validation to a column with Aspose.Cells in C# | save an Aspose.Cells workbook as a JSON file in .NET | apply integer validation to a range of cells and export to JSON using Aspose.Cells | Aspose.Cells C# example for data validation and JSON export | C# code to create workbook, set validation between 1 and 100, and save as JSON
// Tags: Aspose.Cells data validation API C# | Aspose.Cells JSON save format .NET | Aspose.Cells validation range B2:B10 | Aspose.Cells integer range validation | Aspose.Cells workbook export to JSON

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, populates cells A1‑A4 with sample data, defines a whole‑number validation (1‑100) for the range B2:B10 with custom input and error messages, and then saves the workbook to an output.json file using the JSON save format, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["A2"].PutValue("Apple");
            cells["A3"].PutValue("Banana");
            cells["A4"].PutValue("Cherry");

            // Define the cell range B2:B10 for integer validation
            ValidationCollection validations = sheet.Validations;
            CellArea area = new CellArea
            {
                StartRow = 1,   // B2 (zero‑based)
                StartColumn = 1,
                EndRow = 9,     // B10
                EndColumn = 1
            };

            // Add validation and retrieve the Validation object
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];

            // Configure validation as whole number between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";
            validation.InputMessage = "Enter quantity (1-100)";
            validation.ErrorMessage = "Invalid quantity. Must be between 1 and 100.";
            // Optional: show messages (available in newer versions)
            // validation.ShowInputMessage = true;
            // validation.ShowError = true;

            // Save the workbook as JSON
            string outputPath = "output.json";
            workbook.Save(outputPath, SaveFormat.Json);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
