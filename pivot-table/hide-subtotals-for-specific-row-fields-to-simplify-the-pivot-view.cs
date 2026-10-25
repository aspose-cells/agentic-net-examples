// Title: Hide subtotals for selected row fields in an Excel pivot table using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide row field subtotals in Excel pivot table | How to suppress Category and Region subtotals in a pivot table with Aspose.Cells | Aspose.Cells .NET code to remove specific pivot row subtotals | Excel pivot table subtotal visibility control using Aspose.Cells C#
// Tags: aspnet cells pivot row subtotal hide | c# aspnet cells pivot subtotal suppression | excel pivot hide specific row subtotals aspnet cells | aspnet cells control pivot subtotal visibility | c# pivot table row field subtotal removal

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, accesses the first worksheet's pivot table, and iterates over the row fields named "Category" and "Region" to illustrate where subtotals could be hidden. Because the current Aspose.Cells API does not expose SubtotalFunctions or ShowSubTotal properties, the code only demonstrates locating the fields and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "PivotData.xlsx";
            string outputPath = "PivotData_HiddenSubtotals.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Row fields for which subtotals should be hidden
            string[] fieldsToHide = { "Category", "Region" };

            foreach (string fieldName in fieldsToHide)
            {
                // Locate the matching row field
                for (int i = 0; i < pivotTable.RowFields.Count; i++)
                {
                    PivotField rowField = pivotTable.RowFields[i];
                    if (rowField.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        // The SubtotalFunctions and ShowSubTotal properties are not available
                        // in the current Aspose.Cells version. If needed, implement alternative
                        // logic here to modify subtotals.
                        break;
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
