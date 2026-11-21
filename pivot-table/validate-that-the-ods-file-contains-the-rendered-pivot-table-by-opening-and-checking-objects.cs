// Title: Detect rendered pivot tables in an ODS workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to determine if an ODS file has any pivot tables | how to programmatically verify pivot table presence in OpenDocument Spreadsheet using Aspose.Cells | checking for rendered pivot tables in ODS workbooks with Aspose.Cells .NET API | sample C# method to validate pivot table existence in an ODS spreadsheet | Aspose.Cells iterate worksheets to find pivot tables in ODS format
// Tags: Aspose.Cells ODS pivot table detection | C# verify pivot table existence in OpenDocument spreadsheet | iterate worksheets for pivot tables using Aspose.Cells | rendered pivot table validation in ODS files | load ODS workbook and check PivotTables collection

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable classes

// The provided C# example uses Aspose.Cells to load an ODS workbook, iterates through each worksheet, and checks the PivotTables collection. It returns true when a pivot table with at least one row, column, or data field is found, confirming that the ODS file contains a rendered pivot table.
class PivotTableValidator
{
    // Validates that the specified ODS file contains at least one rendered pivot table.
    // Returns true if a pivot table is found; otherwise false.
    public static bool ContainsRenderedPivotTable(string odsFilePath)
    {
        // Ensure the file exists before attempting to load it.
        if (!File.Exists(odsFilePath))
        {
            Console.WriteLine($"File not found: {odsFilePath}");
            return false;
        }

        try
        {
            // Load the ODS workbook.
            Workbook workbook = new Workbook(odsFilePath);

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // The PivotTables collection holds all pivot tables on the worksheet.
                if (sheet.PivotTables.Count > 0)
                {
                    // Examine the first pivot table; if it has any fields defined, consider it rendered.
                    PivotTable pt = sheet.PivotTables[0];
                    if (pt != null &&
                        (pt.RowFields.Count > 0 ||
                         pt.ColumnFields.Count > 0 ||
                         pt.DataFields.Count > 0))
                    {
                        return true;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors during processing.
            Console.WriteLine($"Error processing file '{odsFilePath}': {ex.Message}");
        }

        // No rendered pivot tables found.
        return false;
    }

    // Example usage.
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the ODS file as an argument.");
            return;
        }

        string odsPath = args[0];
        bool hasPivot = ContainsRenderedPivotTable(odsPath);
        Console.WriteLine(hasPivot
            ? "The ODS file contains a rendered pivot table."
            : "No rendered pivot table was found in the ODS file.");
    }
}
