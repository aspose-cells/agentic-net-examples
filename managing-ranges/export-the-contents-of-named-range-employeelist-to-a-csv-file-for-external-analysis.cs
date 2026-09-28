// Title: Export the 'EmployeeList' named range from an Excel workbook to a CSV file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook, finds the named range "EmployeeList", copies its cells to a new workbook, and saves the result as a CSV file with Aspose.Cells. | Show how to check for the existence of a named range before exporting its data to a CSV file using Aspose.Cells in a .NET application. | Demonstrate copying a specific named range to a temporary worksheet and exporting it as CSV without altering the original workbook in C#.
// Common Searches: Aspose.Cells C# export named range to CSV file | How to save a specific named range as CSV using Aspose.Cells .NET | Copy Excel named range to new workbook and export as CSV with Aspose in C# | Verify named range exists before exporting to CSV Aspose.Cells | Export EmployeeList range from Excel to CSV using Aspose.Cells
// Tags: export named range to csv Aspose.Cells | copy range to new workbook C# | verify named range existence Aspose.Cells | save workbook as csv Aspose.Cells | named range EmployeeList extraction

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads 'input.xlsx', retrieves the 'EmployeeList' named range, copies its cells to a fresh workbook, and saves the data as 'EmployeeList.csv' using Aspose.Cells for .NET.
class ExportNamedRangeToCsv
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "EmployeeList.csv";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the named range "EmployeeList"
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Retrieve the named range object
            Name employeeRange = sourceWorkbook.Worksheets.Names["EmployeeList"];
            if (employeeRange == null)
            {
                Console.WriteLine("Named range 'EmployeeList' not found.");
                return;
            }

            // Get the actual range referenced by the name
            AsposeRange range = employeeRange.GetRange();

            // Determine the range boundaries
            int firstRow = range.FirstRow;
            int firstColumn = range.FirstColumn;
            int totalRows = range.RowCount;
            int totalColumns = range.ColumnCount;

            // Get the worksheet that contains the range
            Worksheet sourceSheet = range.Worksheet;

            // Create a new workbook that will hold only the data from the named range
            Workbook csvWorkbook = new Workbook();
            Worksheet destSheet = csvWorkbook.Worksheets[0];

            // Create source and destination Range objects
            AsposeRange srcRange = sourceSheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
            AsposeRange destRange = destSheet.Cells.CreateRange(0, 0, totalRows, totalColumns);

            // Copy the data (values, formulas, formats) from the source range to the destination range
            destRange.Copy(srcRange);

            // Save the new workbook as a CSV file
            csvWorkbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Named range exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
