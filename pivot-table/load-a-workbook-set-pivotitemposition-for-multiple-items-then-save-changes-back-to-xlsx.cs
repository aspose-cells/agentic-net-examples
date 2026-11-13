// Title: Reorder Pivot Table Row Items by Setting PivotItem.Position and Save Workbook to XLSX with Aspose.Cells for .NET
// Common Searches: how to change the order of row field items in an existing Excel pivot table using Aspose.Cells C# | Aspose.Cells set PivotItem.Position for multiple items in a pivot table | C# code to reorder pivot items and save the workbook as XLSX with Aspose.Cells | update pivot table item positions programmatically in .NET with Aspose.Cells | load workbook, modify pivot row items order, and save using Aspose.Cells for .NET
// Tags: Aspose.Cells pivot item repositioning | C# set PivotItem.Position | reorder pivot row field items Aspose.Cells | save workbook to XLSX after pivot update | programmatic pivot table item ordering .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable, PivotField, PivotItem

// The example loads an existing XLSX workbook, checks for a pivot table on the first worksheet, accesses the first row field, reassigns the Position property of up to three PivotItem objects to reorder them, and then saves the modified workbook to a new XLSX file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("The pivot table does not contain any row fields.");
                return;
            }

            // Choose the first row field whose items will be reordered
            PivotField pivotField = pivotTable.RowFields[0];

            // Reorder the first few items (up to three) within the selected pivot field
            int itemsToReorder = Math.Min(3, pivotField.PivotItems.Count);
            for (int i = 0; i < itemsToReorder; i++)
            {
                PivotItem item = pivotField.PivotItems[i];
                item.Position = i; // assign the desired position
            }

            // Save the workbook with the updated pivot item positions
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
