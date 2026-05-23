using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

class ExportNamedRangeToCsv
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "EmployeeList.csv";
        const string namedRange = "EmployeeList";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the named range
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range (use GetRangeByName)
            Aspose.Cells.Range employeeRange = workbook.Worksheets.GetRangeByName(namedRange);
            if (employeeRange == null)
            {
                Console.WriteLine($"Error: Named range \"{namedRange}\" not found.");
                return;
            }

            // Determine the cell area that corresponds to the named range
            int startRow = employeeRange.FirstRow;
            int endRow = startRow + employeeRange.RowCount - 1;
            int startColumn = employeeRange.FirstColumn;
            int endColumn = startColumn + employeeRange.ColumnCount - 1;

            // Configure TxtSaveOptions to export only the specified area as CSV
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                ExportArea = new CellArea
                {
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = startColumn,
                    EndColumn = endColumn
                },
                Separator = ',' // Optional: set a delimiter (comma is default for CSV)
            };

            // Save the selected range to a CSV file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Named range \"{namedRange}\" exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}