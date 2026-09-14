// Title: Adjust slicer item row height to 30 points for a table column with Aspose.Cells for .NET (C#)
// AI Prompts: Create a slicer linked to the second column of a ListObject and set its RowHeight property to 30 points using Aspose.Cells in C#. | Programmatically modify the row height of each slicer item after binding the slicer to a specific table column with Aspose.Cells for .NET.
// Common Searches: how to set slicer row height to 30 points using Aspose.Cells C# | Aspose.Cells example linking slicer to table column and changing item height | C# code to adjust Excel slicer item height with Aspose.Cells | set RowHeight property of slicer items in Aspose.Cells workbook | change slicer item size after linking to ListObject column Aspose.Cells
// Tags: Aspose.Cells slicer RowHeight API usage | C# associate slicer with ListObject column | programmatic adjustment of slicer item dimensions | Excel slicer styling via Aspose.Cells | customize slicer item height C#

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// The example creates a workbook, adds a two‑column table, inserts a slicer linked to the second column, sets each slicer item's row height to 30 points via the RowHeight property, and saves the file as SlicerRowHeight30.xlsx.
class AdjustSlicerRowHeight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a table with two columns
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue("CatA");
        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue("CatB");
        sheet.Cells["A4"].PutValue("Item3");
        sheet.Cells["B4"].PutValue("CatA");

        // Add a ListObject (table) covering the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Add a slicer linked to the second column (index 1) of the table
        // Position the slicer at row 6, column 2 (cell B6)
        int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[1], 5, 1);
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Set the row height of each slicer item to 30 points
        slicer.RowHeight = 30;

        // Save the workbook
        workbook.Save("SlicerRowHeight30.xlsx");
    }
}
