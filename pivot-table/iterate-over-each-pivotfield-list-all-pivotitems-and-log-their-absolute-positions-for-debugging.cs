// Title: How to enumerate all PivotField items and log their zero‑based positions using Aspose.Cells for .NET
// Common Searches: asp.net list pivot field items with their index using Aspose.Cells | c# get absolute position of each pivot item in an Excel pivot table Aspose.Cells | debug pivot table item order with Aspose.Cells .NET | iterate through pivot fields and output item positions Aspose.Cells | retrieve pivot item index programmatically in C# Aspose.Cells
// Tags: Aspose.Cells pivot item enumeration | Aspose.Cells pivot item position logging | C# pivot table field iteration | Excel pivot table debugging Aspose | Aspose.Cells zero‑based item index

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook (creating a new one if the source file is missing), walks through every worksheet and each PivotTable it contains, and for each Row, Column, Data, and Page field calls a helper method that writes the field name and, for every PivotItem in that field, prints the item's name together with its zero‑based Position value for debugging. The workbook is then saved.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Ensure the input file exists; if not, create an empty workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables on the worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    Console.WriteLine($"PivotTable: {pivotTable.Name} (Sheet: {sheet.Name})");

                    // Row fields
                    foreach (PivotField field in pivotTable.RowFields)
                    {
                        LogPivotFieldItems(field);
                    }

                    // Column fields
                    foreach (PivotField field in pivotTable.ColumnFields)
                    {
                        LogPivotFieldItems(field);
                    }

                    // Data fields
                    foreach (PivotField field in pivotTable.DataFields)
                    {
                        LogPivotFieldItems(field);
                    }

                    // Page fields (filters)
                    foreach (PivotField field in pivotTable.PageFields)
                    {
                        LogPivotFieldItems(field);
                    }
                }
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to list items of a pivot field and log their absolute positions
    static void LogPivotFieldItems(PivotField field)
    {
        Console.WriteLine($"  Field: {field.Name}");

        // Iterate over each PivotItem in the field
        foreach (PivotItem item in field.PivotItems)
        {
            // Position is the zero‑based index of the item within the field
            int position = item.Position;

            // Log the item name and its absolute position for debugging
            Console.WriteLine($"    Item: {item.Name}, Position: {position}");
        }
    }
}
