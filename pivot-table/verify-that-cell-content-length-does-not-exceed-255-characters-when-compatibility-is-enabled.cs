// Title: Check and enforce the 255‑character string limit in Excel cells with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that iterates through all worksheets and throws an InvalidOperationException when a string cell contains more than 255 characters. | Modify the provided Aspose.Cells example to collect the addresses of cells exceeding 255 characters and write them to a log file instead of raising an exception. | Create a C# routine with Aspose.Cells that truncates any string longer than 255 characters to the maximum allowed length before saving the workbook.
// Common Searches: asp.net aspose.cells verify cell text length does not exceed 255 characters | how to detect Excel compatibility mode string length violation using Aspose.Cells C# | c# iterate over workbook cells and log cells with text longer than 255 characters Aspose.Cells | throw exception for Excel cell content over 255 characters with Aspose.Cells .NET
// Tags: validate string cell length Aspose.Cells | enforce 255‑character limit Excel compatibility | scan workbook cells for length overflow C# | exception on long cell content Aspose.Cells | truncate oversized cell text Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook with Aspose.Cells, walks through each worksheet's used range, checks every string cell for a length greater than 255 characters, and throws an InvalidOperationException identifying the offending cell before saving the file.
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
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In recent Aspose.Cells versions the compatibility mode property may be unavailable.
            // The 255‑character limit is enforced manually below, so we omit setting IsCompatibilityMode.

            // Iterate through all worksheets and their used cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;
                            if (!string.IsNullOrEmpty(text) && text.Length > 255)
                            {
                                // Cell content exceeds the allowed length; handle as needed
                                throw new InvalidOperationException(
                                    $"Cell {cell.Name} contains {text.Length} characters, exceeding the 255‑character limit.");
                            }
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook (optional, replace with desired output path)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
