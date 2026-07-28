// Title: Export a Named Range to CSV with Aspose.Cells for .NET (C# Sample)
// Description: C# example that loads an .xlsx workbook, retrieves the "EmployeeList" named range, defines its cell area, configures TxtSaveOptions with a comma separator, and saves only that range to a CSV file. Includes checks for missing files and undefined ranges.
// Keywords: Aspose.Cells | C# | .NET | CSV export | named range | EmployeeList | TxtSaveOptions | ExportArea | Excel to CSV | code sample | GitHub example
// Common Searches: Aspose.Cells export named range to CSV C# | How to save a specific Excel range as CSV using Aspose.Cells | TxtSaveOptions ExportArea example | C# code to extract EmployeeList range to CSV | Aspose.Cells CSV export for a single range
// Developer Intent: Generate a CSV file that contains only the data from the "EmployeeList" named range in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create a lightweight CSV report of employee records for external analytics. | Provide a data feed to payroll or HR systems that require only the employee list. | Extract a specific worksheet region for auditing without exposing the full workbook.
// AI Prompts: Write C# code with Aspose.Cells to export the "EmployeeList" named range to a CSV file using a comma separator. | Explain how TxtSaveOptions.ExportArea restricts CSV output to a defined range in Aspose.Cells and show a code snippet. | Add robust error handling for missing workbook files or undefined named ranges when exporting to CSV with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// C# example that loads an .xlsx workbook, retrieves the "EmployeeList" named range, defines its cell area, configures TxtSaveOptions with a comma separator, and saves only that range to a CSV file. Includes checks for missing files and undefined ranges.
class ExportNamedRangeToCsv
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputWorkbook.xlsx";
            const string outputFile = "EmployeeList.csv";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook that contains the named range "EmployeeList"
            Workbook workbook = new Workbook(inputFile);

            // Retrieve the named range object via the Name collection
            Name namedRange = workbook.Worksheets.Names["EmployeeList"];
            if (namedRange == null)
            {
                Console.WriteLine("Error: Named range 'EmployeeList' does not exist in the workbook.");
                return;
            }

            // Get the actual cell range represented by the named range
            Aspose.Cells.Range employeeRange = namedRange.GetRange();

            // Determine the exact cell area of the named range
            int startRow = employeeRange.FirstRow;
            int startColumn = employeeRange.FirstColumn;
            int endRow = startRow + employeeRange.RowCount - 1;
            int endColumn = startColumn + employeeRange.ColumnCount - 1;

            // Configure TxtSaveOptions for CSV export (comma separator) limited to the range area
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ',', // CSV separator
                ExportArea = new CellArea
                {
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = startColumn,
                    EndColumn = endColumn
                }
            };

            // Save the selected range as a CSV file
            workbook.Save(outputFile, saveOptions);
            Console.WriteLine($"Named range exported successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
