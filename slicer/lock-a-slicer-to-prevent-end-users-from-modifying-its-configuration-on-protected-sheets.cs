// Title: How to lock a slicer on a protected worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a table, adds a slicer linked to its first column, locks the slicer's shape and position, and then protects the worksheet with a password using Aspose.Cells. | Show how to prevent end users from moving, resizing, or editing a slicer after the worksheet is protected with Aspose.Cells in a .NET application.
// Common Searches: asp.net lock slicer after worksheet protection using Aspose.Cells | c# Aspose.Cells prevent slicer resizing on protected sheet | how to lock slicer shape in Excel with Aspose.Cells C# | add slicer to listobject and lock it Aspose.Cells example | protect worksheet with password while keeping slicer locked Aspose.Cells
// Tags: lock slicer shape Aspose.Cells | protect worksheet with slicer locked C# | slicer locked position Aspose.Cells | add slicer to listobject Aspose.Cells | worksheet protection password Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerLockDemo
{
    // The example creates a new workbook, defines a table with sample data, inserts a slicer linked to the table's first column, locks the slicer's shape and its position, protects the worksheet with a password, and saves the file as LockedSlicerDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("A");

            // Add a ListObject (table) covering the data range
            int tableIndex = worksheet.ListObjects.Add("A1", "A4", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Add a slicer linked to the first column of the table
            int slicerIndex = worksheet.Slicers.Add(table, 0, "C1");
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Lock the slicer shape so it cannot be modified when the sheet is protected
            slicer.Shape.IsLocked = true;

            // Prevent the slicer from being moved or resized via the UI
            slicer.LockedPosition = true;

            // Protect the worksheet (all protection options) with a password
            worksheet.Protect(ProtectionType.All, "password123", null);

            // Save the workbook
            workbook.Save("LockedSlicerDemo.xlsx");
        }
    }
}
