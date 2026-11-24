// Title: Read the absolute Position of a PivotItem in a row field with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get pivot item absolute position from row field | retrieve PivotItem.Position property in an Excel pivot table using Aspose.Cells | C# code to read the index of the first row field item in a pivot table | how to access pivot item order with Aspose.Cells .NET library | read row field pivot item position in an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells pivot item position retrieval | C# read PivotItem.Position | Aspose.Cells row field item index | Excel pivot table absolute position .NET | Aspose.Cells get pivot item order

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, checks for a worksheet, pivot table, row field, and pivot item, then accesses the first PivotItem of the first row field, reads its Position property (the absolute index), prints the item name and position, and saves the workbook.
class Program
{
    static void Main(string[] args)
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure worksheet has at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("The worksheet contains no pivot tables.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("The pivot table contains no row fields.");
                return;
            }

            PivotField rowField = pivotTable.RowFields[0];

            // Ensure row field has at least one pivot item
            if (rowField.PivotItems.Count == 0)
            {
                Console.WriteLine("The row field contains no pivot items.");
                return;
            }

            PivotItem pivotItem = rowField.PivotItems[0];

            // Retrieve the absolute position of the pivot item
            int absolutePosition = pivotItem.Position;

            Console.WriteLine($"PivotItem '{pivotItem.Name}' absolute position: {absolutePosition}");

            // Save the workbook (even if unchanged)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
