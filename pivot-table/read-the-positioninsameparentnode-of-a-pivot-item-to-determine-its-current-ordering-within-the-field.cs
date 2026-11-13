// Title: Read PivotItem.PositionInSameParentNode to get row field item order with Aspose.Cells for .NET
// Common Searches: how to read the order of pivot row items using Aspose.Cells C# | Aspose.Cells get PositionInSameParentNode of each PivotItem in a pivot table | determine pivot item sequence in Excel row field with Aspose.Cells .NET | C# iterate pivot items and retrieve their position in same parent node | retrieve pivot item index in row field using Aspose.Cells API
// Tags: PivotItem.PositionInSameParentNode Aspose.Cells | read pivot item order .NET | iterate pivot items C# | Excel pivot table row field ordering | Aspose.Cells pivot table API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, accesses the first pivot table's first row field, iterates through its PivotItems, reads each item's PositionInSameParentNode to determine its current order, prints the item name and position, and saves the workbook.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Error: The pivot table does not contain any row fields.");
                return;
            }

            // Choose the first row field
            PivotField pivotField = pivotTable.RowFields[0];

            // Iterate through each pivot item in the selected field
            foreach (PivotItem item in pivotField.PivotItems)
            {
                // PositionInSameParentNode returns the current ordering of the item
                int position = item.PositionInSameParentNode;
                Console.WriteLine($"Item: {item.Name}, Position: {position}");
            }

            // Save the workbook (optional, if modifications were made)
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
