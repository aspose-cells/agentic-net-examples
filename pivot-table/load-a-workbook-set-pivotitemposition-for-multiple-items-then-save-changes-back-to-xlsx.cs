// Title: C# – Reorder Pivot Table Row Items with PivotItem.Position in Aspose.Cells
// Description: Loads an existing XLSX workbook, accesses the first worksheet and its first pivot table, selects the first row field, and changes the Position of specific PivotItems (e.g., "Apple" and "Banana") to reorder them. The pivot table is refreshed, all pivot tables on the sheet are updated, and the workbook is saved with the new item order.
// Keywords: Aspose.Cells C# pivot item position | PivotItem.Position .NET | reorder pivot table rows programmatically | change pivot item order Aspose.Cells | refresh pivot tables after position change | load and save XLSX Aspose.Cells | C# pivot table manipulation | Aspose.Cells example pivot table
// Common Searches: how to move pivot row items to top using Aspose.Cells | Aspose.Cells set PivotItem.Position multiple items | C# reorder items in existing pivot table XLSX | refresh pivot tables after changing item order Aspose | Aspose.Cells pivot table row field reorder code
// Developer Intent: Programmatically reorder specific row items in a pivot table and persist the changes to the workbook.
// Use Cases: Place high‑priority categories (e.g., "Apple") at the beginning of a sales pivot report. | Customize product display order in a financial analysis pivot before exporting. | Batch‑update the order of several row items across multiple pivot tables in a worksheet.
// AI Prompts: Generate C# code that uses Aspose.Cells to set PivotItem.Position for a list of item names and refreshes the pivot table. | Explain the effect of PivotItem.Position in Aspose.Cells and how to verify the new order after saving the workbook. | Provide a C# example that iterates through all row fields of a pivot table and assigns Position values to sort items alphabetically.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    // Loads an existing XLSX workbook, accesses the first worksheet and its first pivot table, selects the first row field, and changes the Position of specific PivotItems (e.g., "Apple" and "Banana") to reorder them. The pivot table is refreshed, all pivot tables on the sheet are updated, and the workbook is saved with the new item order.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                const string outputPath = "OutputWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains a pivot table
                Workbook workbook = new Workbook(inputPath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                // Assume the pivot table is on the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Verify that the worksheet contains at least one pivot table
                if (sheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Access the first pivot table
                PivotTable pivotTable = sheet.PivotTables[0];

                // Ensure the pivot table has at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("The pivot table does not contain any row fields.");
                    return;
                }

                // Choose the first row field whose items we want to reorder
                PivotField rowField = pivotTable.RowFields[0];

                // Access the collection of pivot items for that field
                PivotItemCollection items = rowField.PivotItems;

                // Example: set the Position property for specific items by name.
                // Position specifies the index among all pivot items (global order).
                if (items["Apple"] != null)
                    items["Apple"].Position = 0;   // Move "Apple" to the first position

                if (items["Banana"] != null)
                    items["Banana"].Position = 1; // Move "Banana" to the second position

                // Refresh the pivot table so the changes take effect
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Optionally refresh all pivot tables in the worksheet
                sheet.RefreshPivotTables();

                // Save the workbook with the updated pivot item order
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
