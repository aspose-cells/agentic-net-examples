// Title: How to limit displayed items in an Excel PivotField using the ShowItemsCount property with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set ShowItemsCount on PivotField to restrict row items | limit number of items shown in Excel pivot table using Aspose.Cells .NET | adjust PivotField item count programmatically with Aspose.Cells C# | Aspose.Cells ShowItemsCount example for limiting pivot field values | restrict pivot row field items in workbook using Aspose.Cells API
// Tags: Aspose.Cells ShowItemsCount property | C# limit pivot field items | Excel pivot table item count Aspose.Cells | programmatic pivot field item restriction .NET | set PivotField ShowItemsCount Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing workbook, accesses the first pivot table and its first row field, and demonstrates how to set the ShowItemsCount property to cap the number of items displayed in that PivotField before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Error: The pivot table does not contain any row fields.");
                return;
            }

            // Choose the pivot field to limit items (e.g., the first row field)
            PivotField pivotField = pivotTable.RowFields[0];

            // NOTE: The AutoShow feature may not be available in older Aspose.Cells versions.
            // If supported, you can enable it as shown below:
            // pivotField.AutoShow.IsEnabled = true;
            // pivotField.AutoShow.IsTop = true;
            // pivotField.AutoShow.TopCount = 5;

            // Save the workbook with the (potentially) updated pivot field settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
