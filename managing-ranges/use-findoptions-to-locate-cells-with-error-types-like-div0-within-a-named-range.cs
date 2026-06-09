using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class FindErrorCellsInNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add formulas that will produce #DIV/0! errors
            worksheet.Cells["A1"].Formula = "=1/0";          // #DIV/0!
            worksheet.Cells["A2"].Formula = "=B2";          // No error (B2 is empty)
            worksheet.Cells["B1"].Formula = "=SUM(A1:A2)";  // Propagates the error
            worksheet.Cells["C1"].PutValue("Normal text"); // No error

            // Calculate formulas so that error values are materialized
            workbook.CalculateFormula();

            // Define a named range that includes the cells to be searched
            AsposeRange namedRange = worksheet.Cells.CreateRange("A1", "C2");
            namedRange.Name = "ErrorRange";

            // Retrieve the range object by its name (use fully qualified name to avoid ambiguity)
            AsposeRange errorRange = workbook.Worksheets.GetRangeByName("ErrorRange");

            // Configure FindOptions to search within the named range for the exact error string
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,          // Search cell values (including error strings)
                LookAtType = LookAtType.EntireContent,   // Exact match
                CaseSensitive = false
            };

            // Convert the named range to a CellArea and assign it to FindOptions
            CellArea searchArea = new CellArea
            {
                StartRow = errorRange.FirstRow,
                StartColumn = errorRange.FirstColumn,
                EndRow = errorRange.FirstRow + errorRange.RowCount - 1,
                EndColumn = errorRange.FirstColumn + errorRange.ColumnCount - 1
            };
            findOptions.SetRange(searchArea);

            // The error string representation for division by zero
            const string divZeroError = "#DIV/0!";

            // Iterate to find all cells containing the error within the range
            Cell previousCell = null;
            while (true)
            {
                Cell foundCell = worksheet.Cells.Find(divZeroError, previousCell, findOptions);
                if (foundCell == null)
                    break;

                Console.WriteLine($"Found error cell at: {foundCell.Name}");
                previousCell = foundCell; // Continue searching after the found cell
            }

            // Save the workbook (optional, demonstrates lifecycle compliance)
            string outputPath = "FindErrorCells.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}