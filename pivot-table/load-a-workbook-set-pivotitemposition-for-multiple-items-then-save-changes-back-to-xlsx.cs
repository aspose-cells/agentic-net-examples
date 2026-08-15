// Title: Reorder PivotTable Row Items by Setting PivotItem.Position in Aspose.Cells for .NET
// Description: Loads an XLSX workbook, accesses the first worksheet’s first PivotTable, changes the Position of selected row PivotItems (e.g., "Apple" and "Banana"), refreshes the table, and saves the workbook as a new file.
// Keywords: Aspose.Cells PivotItem.Position | C# reorder pivot table rows | set pivot item order programmatically | Aspose.Cells PivotTable row field | change pivot item sequence .NET | Aspose.Cells workbook save
// Common Searches: Aspose.Cells change pivot item order | Set PivotItem.Position C# | Reorder rows in PivotTable using Aspose.Cells | Move specific pivot items to top Aspose.Cells | Programmatically sort pivot table items .NET
// Developer Intent: Modify the display order of specific row items in a PivotTable and persist the changes to an XLSX file.
// Use Cases: Place "Apple" as the first row entry and "Banana" as the second in a sales‑report PivotTable before exporting. | Align category order in a financial summary PivotTable with a custom hierarchy required by stakeholders. | Enforce a predefined sequence of items in an automatically generated dashboard workbook’s PivotTable.
// AI Prompts: Show C# code that sets PivotItem.Position for multiple items in an Aspose.Cells PivotTable and saves the workbook. | How can I iterate over a list of pivot item names and assign custom positions to them using Aspose.Cells for .NET? | Explain the steps to refresh and recalculate a PivotTable after changing item positions with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an XLSX workbook, accesses the first worksheet’s first PivotTable, changes the Position of selected row PivotItems (e.g., "Apple" and "Banana"), refreshes the table, and saves the workbook as a new file.
class SetPivotItemPositions
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
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

            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            PivotTable pivotTable = sheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Error: The pivot table does not contain any row fields.");
                return;
            }

            // Access the first row field (adjust index if needed)
            PivotField rowField = pivotTable.RowFields[0];
            PivotItemCollection items = rowField.PivotItems;

            // Move specific items to desired global positions
            // Replace "Apple" and "Banana" with actual item names present in your pivot table
            if (items["Apple"] != null)
            {
                items["Apple"].Position = 0; // first position globally
            }

            if (items["Banana"] != null)
            {
                items["Banana"].Position = 1; // second position globally
            }

            // Refresh and recalculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
