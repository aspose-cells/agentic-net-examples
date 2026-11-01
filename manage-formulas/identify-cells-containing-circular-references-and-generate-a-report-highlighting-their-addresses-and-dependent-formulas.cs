// Title: Create a circular reference detection report that lists cell addresses and formulas with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find cells with circular reference errors and export them to a new worksheet | generate a circular reference summary sheet in a .xlsx file using Aspose.Cells | list formula errors and their addresses after calculating workbook with Aspose.Cells | how to programmatically scan used range for error cells in Aspose.Cells .NET | save circular reference report to a separate worksheet in C# Aspose.Cells
// Tags: circular reference detection Aspose.Cells | export formula errors to worksheet C# | scan used range for error cells Aspose.Cells | calculate workbook formulas Aspose.Cells | create summary sheet for circular references .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an XLSX workbook, runs CalculateFormula to surface circular reference errors, iterates each worksheet's used range to collect cells where the value type is IsError, records the cell address and original formula, adds a new worksheet named "CircularReferenceReport" with a header and the collected rows, and saves the updated file.
class CircularReferenceReport
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_with_report.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas; circular references will be marked as errors automatically
            workbook.CalculateFormula();

            // Prepare a list to store report rows (header + data)
            List<string[]> reportRows = new List<string[]>
            {
                new[] { "Cell Address", "Formula" } // header row
            };

            // Scan each worksheet and each used cell for error cells (including circular references)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate only over the used range to improve performance
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        // Check if the cell contains any error (circular references appear as errors)
                        if (cell.Type == CellValueType.IsError)
                        {
                            string address = cell.Name;      // e.g., "B5"
                            string formula = cell.Formula;   // original formula text
                            reportRows.Add(new[] { address, formula });
                        }
                    }
                }
            }

            // Add a new worksheet to hold the report
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "CircularReferenceReport";

            // Write the report data into the new sheet
            Cells reportCells = reportSheet.Cells;
            for (int i = 0; i < reportRows.Count; i++)
            {
                reportCells[i, 0].PutValue(reportRows[i][0]); // Cell Address column
                reportCells[i, 1].PutValue(reportRows[i][1]); // Formula column
            }

            // Save the workbook with the added report sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
