// Title: Validate date strings in column E (yyyy‑MM‑dd) with Aspose.Cells for .NET before saving the workbook
// Common Searches: C# Aspose.Cells validate yyyy-MM-dd dates in a specific Excel column | How to enforce a date format in column E using Aspose.Cells before saving | Throw exception for invalid date cells with Aspose.Cells .NET | Check Excel column for correct date pattern programmatically with Aspose.Cells | Validate Excel worksheet dates format using Aspose.Cells C#
// Tags: Aspose.Cells column E date validation | C# yyyy-MM-dd pattern check in Excel | Excel cell format verification with Aspose.Cells | Invalid date exception handling Aspose.Cells | Workbook save after date validation Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads an Excel file with Aspose.Cells, iterates through all used rows in column E, skips empty cells, and uses DateTime.TryParseExact with the "yyyy-MM-dd" pattern to ensure each cell's text matches the required format. If any cell fails the check, an InvalidOperationException is thrown; otherwise the workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row that contains data
            int maxRow = sheet.Cells.MaxDataRow;

            // Column E (zero‑based index)
            const int columnE = 4;

            // Iterate through each used row in column E
            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = sheet.Cells[row, columnE];

                // Skip null or empty cells
                if (cell.Type == CellValueType.IsNull || string.IsNullOrWhiteSpace(cell.StringValue))
                    continue;

                // Retrieve the cell's displayed text safely
                string cellText = (cell.StringValue ?? string.Empty).Trim();

                // Validate the date format yyyy-MM-dd
                if (!DateTime.TryParseExact(
                        cellText,
                        "yyyy-MM-dd",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out _))
                {
                    throw new InvalidOperationException(
                        $"Cell {cell.Name} does not follow the required date format yyyy-MM-dd.");
                }
            }

            // Save the workbook after successful validation
            workbook.Save(outputPath);
            Console.WriteLine($"Validation succeeded. Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
