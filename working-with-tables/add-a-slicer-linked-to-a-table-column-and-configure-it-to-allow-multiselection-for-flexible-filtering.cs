// Title: Aspose.Cells for .NET – Create a Multi‑Selection Slicer Tied to a Table Column (C#)
// Description: Demonstrates how to generate a workbook, define a ListObject, insert a slicer that references the first column, and configure it for multi‑selection, two‑column layout, movable positioning, and a custom caption before saving as SlicerWithMultiSelection.xlsx.
// Keywords: Aspose.Cells slicer C# | add slicer to ListObject | multi selection slicer .NET | configure slicer properties | table column filter Aspose | Excel slicer programmatic | Workbook with slicer | Aspose.Cells example
// Common Searches: Aspose.Cells add slicer to table column C# | Enable multi‑selection on slicer using Aspose.Cells | Set slicer NumberOfColumns property .NET | Move or resize slicer programmatically Aspose | Save workbook with slicer Aspose.Cells
// Developer Intent: Programmatically attach a slicer to a table column and set it up for flexible multi‑selection filtering.
// Use Cases: Interactive dashboard where users filter sales categories simultaneously. | Dynamic reporting workbook that lets analysts pick multiple product groups. | Customizable Excel view with a movable slicer for on‑the‑fly data exploration.
// AI Prompts: Write C# code with Aspose.Cells to add a slicer linked to the first ListObject column and enable multi‑selection. | Show how to adjust slicer properties such as NumberOfColumns, ShowAllItems, and LockedPosition in Aspose.Cells. | Explain steps to verify that the slicer correctly filters the table when several items are selected.

using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerExample
{
    // Demonstrates how to generate a workbook, define a ListObject, insert a slicer that references the first column, and configure it for multi‑selection, two‑column layout, movable positioning, and a custom caption before saving as SlicerWithMultiSelection.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("A");
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue("B");
            sheet.Cells["B6"].PutValue(50);

            // Add a table that covers the data range
            int tableIndex = sheet.ListObjects.Add(0, 0, 5, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Add a slicer linked to the first column ("Category") of the table
            // Place the slicer starting at row 8, column 1 (cell A8)
            SlicerCollection slicers = sheet.Slicers;
            int slicerIndex = slicers.Add(table, table.ListColumns[0], 8, 1);
            Slicer slicer = slicers[slicerIndex];

            // Configure the slicer for flexible multi‑selection filtering
            slicer.NumberOfColumns = 2;          // Show items in two columns
            slicer.ShowAllItems = true;          // Ensure all items are visible
            slicer.LockedPosition = false;       // Allow the user to move/resize the slicer
            slicer.Caption = "Category Filter";  // Optional caption

            // Save the workbook
            workbook.Save("SlicerWithMultiSelection.xlsx");
        }
    }
}
