// Title: Enumerate and print display names of all pivot fields (row, column, data, page) in every pivot table of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, verifies the file exists, loops through all worksheets and their pivot tables, and writes each pivot field’s DisplayName (row, column, data, page) to the console with proper exception handling. | Create a reusable C# method that accepts a workbook path, loads it via Aspose.Cells, enumerates every pivot table’s fields, and returns or prints the list of field display names.
// Common Searches: Aspose.Cells C# how to get names of all pivot fields in a workbook | list row column data and page field display names from pivot tables using Aspose.Cells | C# iterate over pivot tables in Excel and output field names to console | example code for enumerating pivot fields with Aspose.Cells for .NET
// Tags: Aspose.Cells pivot field enumeration | C# console output Excel pivot field names | Aspose.Cells iterate workbook pivot tables | list pivot row column data page fields Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The example loads an input.xlsx workbook, checks its existence, iterates each worksheet and each pivot table, and writes the DisplayName of every row, column, data, and page field to the console while handling errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table on the worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Output row fields
                    foreach (PivotField field in pivotTable.RowFields)
                    {
                        Console.WriteLine(field.DisplayName);
                    }

                    // Output column fields
                    foreach (PivotField field in pivotTable.ColumnFields)
                    {
                        Console.WriteLine(field.DisplayName);
                    }

                    // Output data fields
                    foreach (PivotField field in pivotTable.DataFields)
                    {
                        Console.WriteLine(field.DisplayName);
                    }

                    // Output page fields
                    foreach (PivotField field in pivotTable.PageFields)
                    {
                        Console.WriteLine(field.DisplayName);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
