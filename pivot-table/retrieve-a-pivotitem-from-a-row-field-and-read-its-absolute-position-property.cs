// Title: C# example: read the absolute Position of the first PivotItem in a row field using Aspose.Cells
// Common Searches: how to get the Position of a PivotItem in a row field with Aspose.Cells C# | Aspose.Cells example for reading pivot item absolute index in .NET | retrieve first row field PivotItem position from an Excel pivot table using C# | C# code to access PivotItem.Position property in Aspose.Cells | read pivot item index from a row field in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells read pivot item position | retrieve pivot item from row field .NET | C# get pivot item absolute index | pivot table row field item Aspose.Cells | Excel pivot item Position property C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an Excel workbook, checks for a pivot table and a row field, retrieves the first PivotItem from that row field, reads its absolute Position property, prints the value, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure a row field exists
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Pivot table has no row fields.");
                return;
            }

            // Get the first row field
            PivotField rowField = pivotTable.RowFields[0];

            // Ensure a pivot item exists
            if (rowField.PivotItems.Count == 0)
            {
                Console.WriteLine("Row field has no pivot items.");
                return;
            }

            // Retrieve the first pivot item
            PivotItem pivotItem = rowField.PivotItems[0];

            // Read the absolute position of the pivot item
            int absolutePosition = pivotItem.Position;

            // Display the position
            Console.WriteLine("PivotItem absolute position: " + absolutePosition);

            // Save the workbook (if any changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
