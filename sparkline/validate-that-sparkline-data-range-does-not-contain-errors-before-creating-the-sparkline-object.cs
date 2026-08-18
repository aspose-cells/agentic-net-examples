// Title: C# – Validate Sparkline Data Range for Errors Before Adding Sparkline Group with Aspose.Cells
// Description: Shows how to scan a worksheet range for Excel error values (e.g., #N/A, #DIV/0!) using Aspose.Cells, then conditionally create a line sparkline group in a .NET workbook. Includes the reusable IsDataRangeValid method, sample data setup, and workbook export.
// Keywords: Aspose.Cells | C# | Sparkline | Data range validation | Error cells | #N/A | #DIV/0 | SparklineGroup | CellArea | .NET | Excel automation | prevent sparkline errors
// Common Searches: Aspose.Cells check error values before sparkline | C# validate sparkline source range | how to avoid #N/A in Aspose.Cells sparkline | prevent sparkline creation when data contains errors | validate cell range for errors Aspose.Cells .NET
// Developer Intent: Ensure a sparkline is only added when its source range has no error values.
// Use Cases: Validate a row of financial metrics before generating a line sparkline in an automated report. | Check a user‑selected dynamic range for #N/A or #DIV/0! before inserting a column sparkline. | Iterate over multiple worksheets, creating sparkline groups only for ranges that pass error‑value validation.
// AI Prompts: Write a C# method for Aspose.Cells that returns false if any cell in a given range contains an error value, then use it to conditionally add a sparkline group. | Generate code that logs the address of the first error cell and skips sparkline creation in Aspose.Cells. | Extend the validation to handle merged cells and custom error handling while creating sparklines with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineValidationExample
{
    // Shows how to scan a worksheet range for Excel error values (e.g., #N/A, #DIV/0!) using Aspose.Cells, then conditionally create a line sparkline group in a .NET workbook. Includes the reusable IsDataRangeValid method, sample data setup, and workbook export.
    class Program
    {
        // Checks whether any cell in the specified range contains an error value.
        static bool IsDataRangeValid(Worksheet sheet, string range)
        {
            // Split the range (e.g., "A1:D1") into start and end addresses.
            string[] parts = range.Split(':');
            if (parts.Length != 2) return false;

            // Convert addresses to CellArea.
            CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);

            // Iterate through each cell in the area.
            for (int row = area.StartRow; row <= area.EndRow; row++)
            {
                for (int col = area.StartColumn; col <= area.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // If the cell type is an error, the range is invalid.
                    if (cell.Type == CellValueType.IsError)
                        return false;
                }
            }
            return true;
        }

        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline.
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(7);
            sheet.Cells["D1"].PutValue(2);
            sheet.Cells["E1"].PutValue(9);

            // Define the data range for the sparkline.
            string dataRange = "A1:E1";

            // Validate the data range before creating the sparkline.
            if (IsDataRangeValid(sheet, dataRange))
            {
                // Define where the sparkline will be placed.
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 5,
                    EndColumn = 5
                };

                // Add a sparkline group using the validated range.
                int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIdx];

                // Optionally add a sparkline explicitly (the Add method already creates one).
                // group.Sparklines.Add(dataRange, 0, 5);
            }
            else
            {
                Console.WriteLine("The specified data range contains errors and cannot be used for a sparkline.");
            }

            // Save the workbook.
            workbook.Save("SparklineValidated.xlsx");
        }
    }
}
