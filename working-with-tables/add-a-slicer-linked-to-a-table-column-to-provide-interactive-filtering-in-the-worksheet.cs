// Title: Add a slicer linked to a table column with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, convert a range to a ListObject, and attach a slicer to the table's Category column. The slicer is positioned, captioned, styled, and the file is saved as SlicerLinkedToTable.xlsx.
// Keywords: Aspose.Cells slicer example | C# add slicer to ListObject | Aspose.Cells interactive filter | table slicer positioning | slicer style Aspose.Cells | Excel slicer .NET | Aspose.Cells GitHub sample | filter dashboard Aspose.Cells
// Common Searches: how to add a slicer to a table using Aspose.Cells C# | Aspose.Cells slicer linked to column example | set slicer caption and style Aspose.Cells | position slicer at specific cell Aspose.Cells | Aspose.Cells interactive filtering tutorial
// Developer Intent: Create a slicer that is bound to a specific table column so users can filter worksheet data interactively.
// Use Cases: Enable finance analysts to filter expense categories in a report with a single click. | Drive a pivot chart that updates automatically when a slicer selects different product lines. | Build a sales dashboard where regional data can be explored via a slicer linked to the Region column.
// AI Prompts: Generate C# code using Aspose.Cells to add a slicer for the 'Region' column of an existing table, display three columns, and apply a dark style. | Show how to modify the caption, style, and cell position of a slicer after it has been added to a worksheet with Aspose.Cells for .NET. | Explain how to read the selected items of a slicer at runtime and use them to filter another ListObject in the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace SlicerExample
{
    // Demonstrates how to create a workbook, convert a range to a ListObject, and attach a slicer to the table's Category column. The slicer is positioned, captioned, styled, and the file is saved as SlicerLinkedToTable.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data for the table
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

            // Convert the range into a ListObject (table)
            // The Add method returns the index of the new table
            int tableIndex = sheet.ListObjects.Add("A1", "B5", true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2; // optional styling

            // Add a slicer linked to the "Category" column of the table
            // Position the slicer at row 7, column 2 (cell B7)
            SlicerCollection slicers = sheet.Slicers;
            int slicerIndex = slicers.Add(table, table.ListColumns[0], 7, 2);
            Slicer slicer = slicers[slicerIndex];

            // Optional: customize slicer appearance
            slicer.Caption = "Category Filter";
            slicer.NumberOfColumns = 1;
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Save the workbook
            workbook.Save("SlicerLinkedToTable.xlsx");
        }
    }
}
