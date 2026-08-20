// Title: Aspose.Cells for .NET – Add a Table Slicer with Multi‑Selection and Custom Layout (C#)
// Description: A complete C# example that creates a workbook, converts a range to a ListObject, inserts a slicer linked to the "Category" column, enables multi‑selection, sets a two‑column layout, shows all items, and allows the slicer to be moved before saving the file.
// Keywords: Aspose.Cells | C# slicer example | table slicer Aspose.Cells | multi‑select slicer | ListObject slicer | Excel filtering with slicer | NumberOfColumns property | ShowAllItems slicer | locked position slicer | GitHub Aspose.Cells sample
// Common Searches: how to add a slicer to a ListObject using Aspose.Cells .NET | Aspose.Cells multi‑select slicer C# example | set slicer NumberOfColumns Aspose.Cells | show all items in Aspose.Cells slicer | move or resize slicer programmatically Aspose.Cells | GitHub Aspose.Cells table slicer demo
// Developer Intent: Insert a slicer linked to a table column and configure it for multi‑selection with a custom layout.
// Use Cases: Enable end‑users to filter a sales summary by multiple categories directly in the workbook. | Create an interactive dashboard where selecting several items updates linked charts and pivot tables. | Provide a reusable method that adds a configurable slicer to any ListObject for dynamic reporting.
// AI Prompts: Write C# code with Aspose.Cells that adds a slicer to a ListObject, enables multi‑selection, sets NumberOfColumns to 3, and locks its position. | Explain the effect of ShowAllItems and NumberOfColumns on slicer behavior in an Aspose.Cells workbook. | Refactor the slicer creation into a method that accepts worksheet, table name, column name, column count, and lock flag.

using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// A complete C# example that creates a workbook, converts a range to a ListObject, inserts a slicer linked to the "Category" column, enables multi‑selection, sets a two‑column layout, shows all items, and allows the slicer to be moved before saving the file.
class SlicerMultiSelectDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will become a table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Drink");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Food");
        sheet.Cells["B4"].PutValue(150);
        sheet.Cells["A5"].PutValue("Snack");
        sheet.Cells["B5"].PutValue(60);

        // Add a ListObject (table) that covers the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B5", true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Add a slicer linked to the "Category" column of the table
        // Use the overload Add(ListObject, ListColumn, int row, int column)
        // The slicer will be placed starting at cell D2 (row index 1, column index 3)
        SlicerCollection slicers = sheet.Slicers;
        int slicerIndex = slicers.Add(table, table.ListColumns[0], 1, 3);
        Slicer slicer = slicers[slicerIndex];

        // Configure slicer to allow flexible multi‑selection
        // In a table slicer multi‑selection is enabled by default,
        // but we can adjust visual layout and ensure all items are shown.
        slicer.NumberOfColumns = 2;          // Show items in two columns for easier selection
        slicer.ShowAllItems = true;          // Ensure all items appear even if they have no data
        slicer.LockedPosition = false;       // Allow the user to move/resize the slicer

        // Save the workbook
        workbook.Save("SlicerMultiSelectDemo.xlsx");
    }
}
