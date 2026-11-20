// Title: Detect and log null or empty cells in an Excel worksheet and generate a report with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find cells with null or blank values in a worksheet | how to create a report of empty cells using Aspose.Cells .NET | enumerate all cells in used range and capture blank cell addresses with Aspose.Cells | C# Aspose.Cells detect missing data and export cell addresses to a new workbook | handle password protected Excel file while scanning for null cells with Aspose.Cells
// Tags: null cell detection Aspose.Cells | used range enumeration Aspose.Cells C# | blank cell address export to Excel | null cells report generation Aspose.Cells | password protected workbook handling Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates over every cell in the first worksheet's used range, identifies cells whose Value is null or whose StringValue is empty, records their addresses, creates a new workbook listing those addresses, saves the report, and includes error handling for missing files and password‑protected workbooks.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string reportPath = "NullCellsReport.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied; will throw if protected)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Collect addresses of cells that contain null or empty values
            List<string> nullCellAddresses = new List<string>();

            // Enumerate all cells in the used range
            foreach (Cell cell in cells)
            {
                // Detect null or empty values
                if (cell.Value == null || string.IsNullOrEmpty(cell.StringValue))
                {
                    Console.WriteLine($"Null/empty cell found at {cell.Name}");
                    nullCellAddresses.Add(cell.Name);
                }
            }

            // Create a simple report workbook listing the null cells
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "NullCellsReport";

            // Header
            reportSheet.Cells[0, 0].PutValue("Cell Address");

            // Populate report
            for (int i = 0; i < nullCellAddresses.Count; i++)
            {
                reportSheet.Cells[i + 1, 0].PutValue(nullCellAddresses[i]);
            }

            // Save the report
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report saved to \"{reportPath}\".");
        }
        catch (CellsException ex) when (ex.Message.Contains("password"))
        {
            Console.WriteLine("Error: The workbook is password protected. Provide a valid password to open it.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
