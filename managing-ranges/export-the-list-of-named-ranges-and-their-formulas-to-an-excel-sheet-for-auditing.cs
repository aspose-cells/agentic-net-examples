// Title: Export all named ranges and their RefersTo formulas to a new worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to list named ranges with their RefersTo formulas using Aspose.Cells | How to create an audit worksheet for named ranges in an Excel file with Aspose.Cells .NET | Export named range definitions to a separate sheet in Aspose.Cells | Retrieve and save workbook named ranges to a new sheet in C# | Aspose.Cells example for writing named range details to a new worksheet
// Tags: write named range data to worksheet Aspose.Cells | retrieve RefersTo formula C# Aspose.Cells | add audit sheet to workbook .NET | enumerate workbook names collection Aspose.Cells | save workbook after inserting audit worksheet

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing Excel workbook, adds a new worksheet named "NamedRangesAudit", writes column headers, iterates through the workbook's named ranges, records each range's name and its RefersTo formula, and saves the updated file—all using Aspose.Cells for .NET with proper exception handling.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook (no password supplied)
            Workbook srcWorkbook = new Workbook(inputPath);

            // Add a new worksheet to hold the audit information
            Worksheet auditSheet = srcWorkbook.Worksheets.Add("NamedRangesAudit");

            // Write column headers
            auditSheet.Cells[0, 0].PutValue("Name");
            auditSheet.Cells[0, 1].PutValue("Refers To");

            // Iterate through all named ranges in the workbook
            int rowIndex = 1;
            foreach (Name namedRange in srcWorkbook.Worksheets.Names)
            {
                auditSheet.Cells[rowIndex, 0].PutValue(namedRange.Text);
                auditSheet.Cells[rowIndex, 1].PutValue(namedRange.RefersTo);
                rowIndex++;
            }

            // Save the workbook with the audit sheet added
            srcWorkbook.Save(outputPath);
            Console.WriteLine($"Audit sheet saved to {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password required)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
