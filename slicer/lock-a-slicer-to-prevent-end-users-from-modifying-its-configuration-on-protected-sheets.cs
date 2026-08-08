// Title: Lock a slicer on a protected worksheet with Aspose.Cells for .NET
// Description: Shows how to create a table‑linked slicer, lock its settings and position, protect the worksheet with a password, and save the workbook using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | lock slicer | slicer protection | read‑only slicer | worksheet protect | Excel slicer lock | IsLocked property | LockedPosition property | Aspose.Cells API
// Common Searches: Aspose.Cells lock slicer | make slicer read only in Excel C# | prevent slicer movement after sheet protection | C# code to protect slicer settings Aspose | disable slicer editing Aspose.Cells
// Developer Intent: Prevent end users from modifying slicer filters, size, or position on a protected Excel sheet.
// Use Cases: Standard reporting templates where slicer choices must stay consistent across users. | Shared dashboards that allow data entry but keep slicer layout fixed. | Excel files distributed to clients with a locked slicer to enforce a uniform view.
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer to a table, locks its configuration and position, and protects the worksheet with a password. | Explain how to lock slicer settings in Aspose.Cells for .NET and note any deprecated properties such as IsLocked. | Provide a step‑by‑step example of protecting a worksheet while keeping slicers read‑only using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to create a table‑linked slicer, lock its settings and position, protect the worksheet with a password, and save the workbook using Aspose.Cells in C#.
    public class LockSlicerDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                // Add a table that will be linked to the slicer
                int tableIndex = worksheet.ListObjects.Add("A1", "A4", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Add a slicer for the first column of the table
                int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Lock the slicer so users cannot change its configuration when the sheet is protected
                slicer.IsLocked = true;          // Prevents editing of slicer settings (property may be obsolete)
                slicer.LockedPosition = true;   // Prevents moving or resizing the slicer

                // Protect the worksheet (all protection options) with a password
                worksheet.Protect(ProtectionType.All, "password123", null);

                // Save the workbook
                workbook.Save("LockedSlicerDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LockSlicerDemo.Run();
        }
    }
}
