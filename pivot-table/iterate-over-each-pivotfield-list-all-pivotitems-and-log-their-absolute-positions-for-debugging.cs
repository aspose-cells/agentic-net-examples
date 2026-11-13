// Title: Iterate PivotFields and log each PivotItem’s absolute Position using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate pivot fields and get item positions | how to retrieve absolute position of PivotItem in Aspose.Cells | debugging pivot table items with Aspose.Cells .NET | list all row and column pivot items and their positions in C# | Aspose.Cells example to print PivotItem.Position for each field
// Tags: pivotfield iteration Aspose.Cells | pivotitem position retrieval C# | enumerate pivot items Aspose.Cells | debug pivot table items .NET | row and column pivotfield processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable, PivotField, PivotItem

// The sample loads an Excel workbook, walks through every worksheet and its pivot tables, iterates over both row and column PivotFields, and writes each PivotItem's name together with its absolute Position value to the console, then optionally saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables on the current worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    Console.WriteLine($"PivotTable: {pivotTable.Name} (Sheet: {sheet.Name})");

                    // Iterate over row fields
                    foreach (PivotField field in pivotTable.RowFields)
                    {
                        Console.WriteLine($"  Row PivotField: {field.Name}");

                        // List all items in the current field and log their absolute positions
                        foreach (PivotItem item in field.PivotItems)
                        {
                            // Position property gives the absolute position of the item within the field
                            Console.WriteLine($"    Item: {item.Name}, Position: {item.Position}");
                        }
                    }

                    // Iterate over column fields (optional, if you also need column fields)
                    foreach (PivotField field in pivotTable.ColumnFields)
                    {
                        Console.WriteLine($"  Column PivotField: {field.Name}");

                        foreach (PivotItem item in field.PivotItems)
                        {
                            Console.WriteLine($"    Item: {item.Name}, Position: {item.Position}");
                        }
                    }
                }
            }

            // Save the workbook if any modifications were made (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
