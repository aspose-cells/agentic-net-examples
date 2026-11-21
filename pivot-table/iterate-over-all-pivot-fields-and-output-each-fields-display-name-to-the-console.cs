// Title: C# example: iterate all pivot table fields in an Excel workbook and print each field’s display name with Aspose.Cells
// Common Searches: how to list every pivot table field name in an .xlsx using Aspose.Cells C# | C# code to enumerate row, column, page and data fields of a pivot table with Aspose.Cells | example for printing pivot field display names from a workbook in .NET | retrieve pivot table field names programmatically in C# Aspose.Cells | Aspose.Cells iterate pivot fields and output to console
// Tags: Aspose.Cells pivot field enumeration | C# list pivot table field display names | iterate pivot fields .NET | read pivot table field names Aspose.Cells | console output pivot field names C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The sample loads an input.xlsx workbook with Aspose.Cells, walks through each worksheet and its pivot tables, and writes the DisplayName of every row, column, page, and data PivotField to the console, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table on the worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Output all pivot fields (row, column, page, data) display names
                    PrintPivotFields(pivotTable);
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to print display names of all fields in a pivot table
    private static void PrintPivotFields(PivotTable pivotTable)
    {
        try
        {
            // Row fields
            foreach (PivotField field in pivotTable.RowFields)
            {
                Console.WriteLine(field.DisplayName);
            }

            // Column fields
            foreach (PivotField field in pivotTable.ColumnFields)
            {
                Console.WriteLine(field.DisplayName);
            }

            // Page fields
            foreach (PivotField field in pivotTable.PageFields)
            {
                Console.WriteLine(field.DisplayName);
            }

            // Data fields
            foreach (PivotField field in pivotTable.DataFields)
            {
                Console.WriteLine(field.DisplayName);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read pivot fields: {ex.Message}");
        }
    }
}
