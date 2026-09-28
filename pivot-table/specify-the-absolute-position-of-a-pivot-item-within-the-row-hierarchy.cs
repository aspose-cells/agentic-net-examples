// Title: Set absolute position of a pivot item in a row hierarchy with Aspose.Cells for .NET (C#)
// AI Prompts: Assign PositionInSameParentNode to a pivot item to move it to a specific index within its row hierarchy using Aspose.Cells in C#. | Reorder pivot table row items by setting their absolute positions programmatically with Aspose.Cells for .NET. | Update the order of specific row field items in an existing pivot table and recalculate the data using C#.
// Common Searches: Aspose.Cells C# set pivot row item order using PositionInSameParentNode | How to change the absolute position of a pivot item in a .NET workbook | Programmatically reorder items in a pivot table row hierarchy with Aspose.Cells | Move pivot table row field items to first position in C# | Aspose.Cells example for positioning pivot items in row hierarchy
// Tags: Aspose.Cells pivot item absolute positioning | C# set pivot row item order | PositionInSameParentNode property Aspose.Cells | reorder pivot table rows .NET | modify pivot hierarchy programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example loads or creates a workbook, builds a pivot table, accesses the first row field's items, and uses the PositionInSameParentNode property to set absolute positions for "ItemA" and "ItemB", thereby reordering them. After recalculating the pivot table, the modified workbook is saved.
    class SetPivotItemAbsolutePosition
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "PivotSource.xlsx";
                Workbook workbook;

                // Load existing workbook or create a sample one if it does not exist
                if (File.Exists(sourcePath))
                {
                    workbook = new Workbook(sourcePath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];

                    // Sample data for the pivot table
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Value");
                    ws.Cells["A2"].PutValue("ItemA");
                    ws.Cells["B2"].PutValue(10);
                    ws.Cells["A3"].PutValue("ItemB");
                    ws.Cells["B3"].PutValue(20);
                    ws.Cells["A4"].PutValue("ItemC");
                    ws.Cells["B4"].PutValue(30);

                    // Create a pivot table based on the sample data
                    string sourceRange = "A1:B4";
                    int pivotRow = 6;
                    int pivotColumn = 0;
                    int pivotIndex = ws.PivotTables.Add(sourceRange, pivotRow, pivotColumn, "PivotTable1");
                    PivotTable pt = ws.PivotTables[pivotIndex];

                    // Add row and data fields using AddFieldToArea
                    pt.AddFieldToArea(PivotFieldType.Row, 0);   // Column A as row field
                    pt.AddFieldToArea(PivotFieldType.Data, 1);  // Column B as data field

                    // Refresh and calculate the pivot table
                    pt.RefreshData();
                    pt.CalculateData();

                    // Save the generated source workbook for future runs
                    workbook.Save(sourcePath);
                }

                // Access the first worksheet and its first pivot table
                Worksheet worksheet = workbook.Worksheets[0];
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                PivotTable pivotTable = worksheet.PivotTables[0];
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("The pivot table does not contain any row fields.");
                    return;
                }

                // Obtain the first row field and its items
                PivotField rowField = pivotTable.RowFields[0];
                PivotItemCollection items = rowField.PivotItems;

                // Retrieve specific items by name using indexer
                PivotItem itemA = items["ItemA"];
                PivotItem itemB = items["ItemB"];

                if (itemA != null && itemB != null)
                {
                    // Set absolute positions within the same parent node
                    itemA.PositionInSameParentNode = 0; // Move "ItemA" to first position
                    itemB.PositionInSameParentNode = 1; // Move "ItemB" to second position
                }
                else
                {
                    Console.WriteLine("Required pivot items ('ItemA' or 'ItemB') were not found.");
                }

                // Recalculate the pivot table to apply changes
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save("PivotModified.xlsx");
                Console.WriteLine("Pivot item positions updated and workbook saved as 'PivotModified.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
