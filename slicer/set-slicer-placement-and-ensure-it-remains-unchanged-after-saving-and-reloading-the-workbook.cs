// Title: How to set slicer position and size in Aspose.Cells for .NET and ensure dimensions persist after saving and reopening the workbook
// AI Prompts: Create a worksheet, add a ListObject, build a pivot table, insert a slicer at cell G2, set its Width to 150 and Height to 200 points, save as XLSX, reload the file, and verify the slicer dimensions are unchanged. | Programmatically define the top‑left cell for an Aspose.Cells slicer, assign specific width and height values, persist the workbook, then read it back to confirm the slicer retains its size. | Adjust slicer placement and size using the Aspose.Cells Slicers.Add method, save the workbook, open it again, and check that the slicer’s Width and Height properties match the original settings.
// Common Searches: aspnet set slicer width height aspose.cells persist after save | c# aspose.cells slicer placement cell G2 example | verify slicer dimensions after reloading Excel workbook using Aspose.Cells | how to keep slicer size unchanged when saving workbook with Aspose.Cells | aspose.cells slicer size persistence across save/load cycle
// Tags: Aspose.Cells slicer size persistence | C# set slicer dimensions Excel | slicer placement cell G2 Aspose.Cells | pivot table slicer configuration .NET | slicer dimension validation after reload

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a new workbook, adds sample data, converts the range to a ListObject, builds a pivot table, inserts a slicer for the 'Category' field at cell G2, sets the slicer's width and height, saves the workbook as XLSX, reloads it, retrieves the slicer, and confirms that its dimensions remain unchanged after the save‑load cycle.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Populate sample data (required for a slicer)
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Convert the range to a table (ListObject) so a slicer can be attached
            int firstRow = 0;          // zero‑based index
            int firstColumn = 0;
            int totalRows = 4;         // includes header row
            int totalColumns = 2;
            int tableIndex = ws.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = ws.ListObjects[tableIndex];
            table.DisplayName = "DataTable";

            // Create a pivot table based on the ListObject (required for slicer in this API version)
            int pivotIndex = ws.PivotTables.Add(table.DisplayName, "E1", "PivotTable1");
            PivotTable pivot = ws.PivotTables[pivotIndex];

            // Add a slicer for the "Category" column of the pivot table.
            // The overload requires the upper‑left cell where the slicer will be placed.
            int slicerIndex = ws.Slicers.Add(pivot, "Category", "G2");
            Slicer slicer = ws.Slicers[slicerIndex];

            // Set slicer size (width and height in points)
            double slicerWidth = 150;   // points
            double slicerHeight = 200;  // points
            slicer.Width = slicerWidth;
            slicer.Height = slicerHeight;

            // Save the workbook
            string filePath = "SlicerDemo.xlsx";
            wb.Save(filePath, SaveFormat.Xlsx);

            // Reload the workbook only if the file exists
            if (File.Exists(filePath))
            {
                try
                {
                    Workbook wbReloaded = new Workbook(filePath);
                    Worksheet wsReloaded = wbReloaded.Worksheets[0];

                    if (wsReloaded.Slicers.Count > 0)
                    {
                        Slicer reloadedSlicer = wsReloaded.Slicers[0];

                        // Verify that size remains unchanged after reload
                        bool sizeUnchanged = Math.Abs(reloadedSlicer.Width - slicerWidth) < 0.01 &&
                                             Math.Abs(reloadedSlicer.Height - slicerHeight) < 0.01;

                        Console.WriteLine("Size unchanged: " + sizeUnchanged);
                    }
                    else
                    {
                        Console.WriteLine("No slicer found after reload.");
                    }
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine("Error loading workbook: " + loadEx.Message);
                }
            }
            else
            {
                Console.WriteLine("Error: The file was not saved correctly.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
