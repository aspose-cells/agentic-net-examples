// Title: Reorder Pivot Table Row Items by Setting PivotItem.Position and Save Workbook to XLSX with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to set PivotItem.Position for multiple items in a pivot table | reorder row field items in an existing Excel pivot table using Aspose.Cells .NET | save modified workbook after changing pivot item order with Aspose.Cells | Aspose.Cells change order of pivot table items programmatically in C# | update pivot table row items position and export to XLSX using Aspose.Cells
// Tags: Aspose.Cells set PivotItem.Position C# | reorder pivot row items Aspose.Cells | modify existing pivot table Aspose.Cells | save workbook as XLSX Aspose.Cells | pivot table item ordering .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an existing XLSX workbook, accesses the first pivot table, reorders the items of its first row field by assigning new values to each PivotItem.Position based on a predefined list, and then saves the updated workbook to a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Pivot table does not contain any row fields.");
                return;
            }

            // Choose the first row field whose items we want to reorder
            PivotField pivotField = pivotTable.RowFields[0];

            // Desired order of items (by their display names)
            List<string> desiredOrder = new List<string> { "ItemB", "ItemA", "ItemC" };

            // Reorder each item by setting its Position property (zero‑based)
            for (int i = 0; i < desiredOrder.Count; i++)
            {
                string itemName = desiredOrder[i];
                PivotItem item = pivotField.PivotItems[itemName];
                if (item != null)
                {
                    item.Position = i;
                }
                else
                {
                    Console.WriteLine($"Pivot item '{itemName}' not found.");
                }
            }

            // Save the workbook with the updated pivot item positions
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
