// Title: How to validate that required worksheets like Invoice and Summary exist in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells check if worksheet exists before processing in C# | C# code to ensure specific sheets are present in an Excel file with Aspose.Cells | Throw exception when a required sheet is missing in an Aspose.Cells workbook | Validate presence of multiple worksheets in a .xlsx using Aspose.Cells .NET | How to verify required worksheets in an Excel workbook loaded with Aspose.Cells
// Tags: worksheet existence check Aspose.Cells | required sheet verification .NET | missing worksheet invalid operation exception | Excel workbook sheet validation C# | Aspose.Cells ensure specific worksheets

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file with Aspose.Cells, iterates over a list of required sheet names (e.g., "Invoice" and "Summary"), and throws an InvalidOperationException if any of those worksheets are absent, ensuring the workbook contains all needed sheets before further processing.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string filePath = "input.xlsx";

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Validate that the required worksheets exist
            ValidateWorksheets(workbook, new[] { "Invoice", "Summary" });

            // Further processing can continue here...
            Console.WriteLine("All required worksheets are present.");
        }
        catch (InvalidOperationException ex)
        {
            // Handle missing worksheet errors
            Console.WriteLine($"Validation error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    static void ValidateWorksheets(Workbook workbook, string[] requiredSheets)
    {
        foreach (string sheetName in requiredSheets)
        {
            // workbook.Worksheets returns null if the sheet is not found
            if (workbook.Worksheets[sheetName] == null)
            {
                throw new InvalidOperationException($"Required worksheet \"{sheetName}\" is missing.");
            }
        }
    }
}
