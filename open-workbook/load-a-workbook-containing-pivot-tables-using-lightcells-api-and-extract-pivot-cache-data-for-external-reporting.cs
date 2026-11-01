// Title: Enumerate row, column, data, and page fields of pivot tables in an XLSX workbook using Aspose.Cells for .NET
// Common Searches: how to list pivot table row and column fields with Aspose.Cells in C# | extract pivot table metadata from an Excel file using Aspose.Cells .NET API | retrieve pivot cache field names for external reporting with Aspose.Cells C# | iterate through all pivot tables in a workbook and get their field lists using Aspose.Cells | C# code to display pivot table page (filter) fields from an XLSX workbook
// Tags: Aspose.Cells pivot table field enumeration C# | read pivot cache metadata .NET | extract pivot row column fields Aspose.Cells | list pivot data page fields C# | process XLSX workbook pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an XLSX workbook with Aspose.Cells, walks through each worksheet and its pivot tables, and prints the pivot table name together with the names of its row, column, data, and page (filter) fields, while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Path to the workbook that contains pivot tables
        string workbookPath = "PivotWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook (default LoadOptions are sufficient for XLSX)
            Workbook workbook = new Workbook(workbookPath);

            // Extract information about each pivot table in the workbook
            ExtractPivotInfo(workbook);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    static void ExtractPivotInfo(Workbook workbook)
    {
        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all pivot tables on the current worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                try
                {
                    Console.WriteLine($"Pivot Table: {pivot.Name} (Worksheet: {sheet.Name})");

                    // Display row fields
                    Console.WriteLine("  Row Fields:");
                    foreach (PivotField field in pivot.RowFields)
                    {
                        Console.WriteLine($"    {field.Name}");
                    }

                    // Display column fields
                    Console.WriteLine("  Column Fields:");
                    foreach (PivotField field in pivot.ColumnFields)
                    {
                        Console.WriteLine($"    {field.Name}");
                    }

                    // Display data fields
                    Console.WriteLine("  Data Fields:");
                    foreach (PivotField field in pivot.DataFields)
                    {
                        Console.WriteLine($"    {field.Name}");
                    }

                    // Display page fields (filter fields)
                    Console.WriteLine("  Page Fields:");
                    foreach (PivotField field in pivot.PageFields)
                    {
                        Console.WriteLine($"    {field.Name}");
                    }

                    Console.WriteLine(); // Blank line between pivot tables
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error processing pivot table '{pivot.Name}': {ex.Message}");
                }
            }
        }
    }
}
