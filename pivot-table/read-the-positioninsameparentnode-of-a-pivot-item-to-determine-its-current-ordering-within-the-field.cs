// Title: Read the PositionInSameParentNode of Pivot Items to Determine Their Order in a PivotTable with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to get the PositionInSameParentNode of a pivot item | retrieve the order of row field items in a PivotTable using Aspose.Cells | list pivot items with their current positions in the same parent node in .NET | read pivot item ranking from a pivot table programmatically with Aspose.Cells | determine pivot item sequence in an Excel file using C# Aspose.Cells API
// Tags: Aspose.Cells pivot item position retrieval | C# pivot table item ordering | PositionInSameParentNode usage Aspose.Cells | row field pivot items enumeration Aspose.Cells | determine pivot item sequence .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook, accesses its first worksheet and pivot table, selects the first row field, iterates through each PivotItem, reads the PositionInSameParentNode property to obtain the item's current order, and prints the item name with its position, while handling missing files, tables, or fields and catching exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Pivot table has no row fields.");
                return;
            }

            // Select the first row field
            PivotField pivotField = pivotTable.RowFields[0];

            // Iterate through each pivot item in the selected field
            foreach (PivotItem item in pivotField.PivotItems)
            {
                int position = item.PositionInSameParentNode;
                Console.WriteLine($"Item: {item.Name}, PositionInSameParentNode: {position}");
            }

            // Uncomment to save any modifications
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
