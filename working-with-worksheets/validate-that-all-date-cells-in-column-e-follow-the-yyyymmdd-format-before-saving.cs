// Title: Validate date cells in column E (yyyy‑mm‑dd) with Aspose.Cells for .NET before saving the workbook
// AI Prompts: Load a workbook using Aspose.Cells, iterate through column E, apply a C# regular expression to ensure each non‑empty cell matches the yyyy‑mm‑dd pattern, and raise an InvalidOperationException when a mismatch occurs. | Insert a pre‑save step that scans every populated cell in column E for the required date format and only invokes Workbook.Save after all cells pass the regex test.
// Common Searches: aspocells c# verify yyyy-mm-dd dates in specific column before saving | how to enforce date format in Excel column using Aspose.Cells .NET | c# iterate column E and validate date strings with Aspose.Cells | throw exception for invalid date format in Excel worksheet Aspose.Cells
// Tags: Aspose.Cells column‑E date validation | C# regex for Excel yyyy‑mm‑dd | date format check before workbook save | InvalidOperationException for Excel date error | enforce yyyy‑mm‑dd pattern in .xlsx

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The program loads "input.xlsx", scans each non‑empty cell in column E of the first worksheet, validates the cell text against a yyyy‑mm‑dd regular expression, throws an InvalidOperationException on any mismatch, ensures the output directory exists, and saves the workbook to "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file \"{inputPath}\" was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Regex to match the required yyyy-mm-dd format
            Regex dateRegex = new Regex(@"^\d{4}-\d{2}-\d{2}$");

            // Determine the last used row in the sheet
            int maxRow = cells.MaxDataRow;

            // Iterate through each cell in column E (zero‑based index 4)
            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = cells[row, 4];

                // Skip empty cells (type IsNull indicates no value)
                if (cell.Type == CellValueType.IsNull)
                    continue;

                // Get the cell's displayed text safely
                string cellText = cell.StringValue?.Trim() ?? string.Empty;

                // Validate the format
                if (!dateRegex.IsMatch(cellText))
                {
                    throw new InvalidOperationException(
                        $"Invalid date format in column E at row {row + 1}: \"{cellText}\". Expected format: yyyy-mm-dd.");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook after successful validation
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
