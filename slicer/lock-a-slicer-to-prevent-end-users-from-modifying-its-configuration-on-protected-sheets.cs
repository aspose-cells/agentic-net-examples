// Title: Lock an Excel slicer on a protected worksheet using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a table, inserts a slicer linked to the table, and then locks the slicer’s configuration and position. The worksheet is protected with a password, preventing end users from moving, resizing, or changing slicer settings while still allowing data filtering.
// Keywords: Aspose.Cells | C# | Excel slicer lock | worksheet protection | IsLocked property | LockedPosition property | slicer shape lock | protect sheet with password | sample code | GitHub example | API reference
// Common Searches: lock slicer Aspose.Cells C# | prevent slicer editing on protected sheet | Aspose.Cells IsLocked and LockedPosition usage | protect worksheet while keeping slicer fixed | C# code to lock Excel slicer with Aspose.Cells
// Developer Intent: Secure a slicer so users cannot modify its settings on a protected worksheet.
// Use Cases: Distribute a dashboard template where slicers stay in a fixed position and retain their filter settings. | Create a reporting workbook that allows filtering but prevents users from moving or resizing slicers. | Automate generation of Excel files with locked slicers to ensure consistent UI across multiple recipients.
// AI Prompts: Show C# code that adds a slicer to a table and locks its configuration on a protected worksheet using Aspose.Cells. | Explain the difference between slicer.IsLocked and slicer.LockedPosition in Aspose.Cells for .NET. | Provide a step‑by‑step example of protecting an Excel sheet with a password while keeping slicer properties immutable.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerLockDemo
{
    // This example creates a workbook, adds a table, inserts a slicer linked to the table, and then locks the slicer’s configuration and position. The worksheet is protected with a password, preventing end users from moving, resizing, or changing slicer settings while still allowing data filtering.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a ListObject (table) covering the data range
            int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Add a slicer linked to the first column of the table
            int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Lock the slicer shape so users cannot change its configuration when the sheet is protected
            slicer.IsLocked = true;               // Prevent editing of slicer settings
            slicer.LockedPosition = true;         // Prevent moving or resizing the slicer

            // Protect the worksheet (all protection options) with a password
            worksheet.Protect(ProtectionType.All, "myPassword", null);

            // Save the workbook
            workbook.Save("SlicerLockedDemo.xlsx");
        }
    }
}
