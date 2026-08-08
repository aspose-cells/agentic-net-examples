// Title: Delete a slicer by name in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a table and pivot table, attach a slicer, then locate the slicer by its assigned name and remove it using the Aspose.Cells API. The example shows the clean‑up of unused slicer controls before saving the file.
// Keywords: Aspose.Cells delete slicer | remove slicer C# | Aspose.Cells slicer API | C# workbook slicer removal | Aspose.Cells .NET pivot slicer | programmatic slicer cleanup | Aspose.Cells remove control
// Common Searches: how to remove a slicer with Aspose.Cells C# | delete slicer by name Aspose.Cells .NET | Aspose.Cells remove unused slicer controls | C# code to delete slicer from workbook | Aspose.Cells slicer removal example
// Developer Intent: Programmatically delete a specific slicer from a worksheet when its name is known.
// Use Cases: Eliminate slicers that were added for temporary analysis before delivering the final report. | Automate workbook cleanup in a reporting pipeline that generates pivot tables with slicers. | Remove slicers that match a naming convention after batch processing multiple workbooks.
// AI Prompts: Generate C# code using Aspose.Cells to delete a slicer by its name from a worksheet. | Show how to enumerate slicer names in a workbook and remove a selected one with Aspose.Cells for .NET. | Explain how to confirm that a slicer has been successfully removed after calling the Remove method.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerRemovalDemo
{
    // Demonstrates how to create a workbook, add a table and pivot table, attach a slicer, then locate the slicer by its assigned name and remove it using the Aspose.Cells API. The example shows the clean‑up of unused slicer controls before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for a table
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("A");
                cells["B4"].PutValue(30);

                // Add a ListObject (table) based on the data range
                int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Add a pivot table that uses the table as its source
                int pivotIndex = sheet.PivotTables.Add("=Sheet1!A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table.
                // The destination cell must be a valid cell address (e.g., "E2").
                int slicerIndex = sheet.Slicers.Add(pivot, "E2", "Category");
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Name = "MySlicer"; // Assign a specific name for later removal

                // ----- Remove the slicer by its name -----
                Slicer slicerToRemove = sheet.Slicers["MySlicer"];
                sheet.Slicers.Remove(slicerToRemove);
                // -----------------------------------------

                // Save the workbook to verify that the slicer has been removed
                workbook.Save("SlicerRemovalResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
