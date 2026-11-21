// Title: How to explicitly dispose Aspose.Cells Workbook objects in C# to free file handles and memory
// Common Searches: C# dispose Aspose.Cells workbook after saving to release file lock | how to free memory used by Aspose.Cells Workbook objects | explicitly close Excel workbook with Aspose.Cells to avoid locked file | using Workbook.Dispose instead of using statement in Aspose.Cells | best practice for releasing resources of Aspose.Cells workbooks in .NET
// Tags: Aspose.Cells workbook disposal pattern | C# release Excel file handle | memory cleanup after Workbook.Save | avoid file lock Aspose.Cells | explicit Workbook.Dispose call | resource cleanup Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample creates a workbook, writes data, saves it, then calls Workbook.Dispose to release file handles and memory. It then loads the saved file, updates cells, saves a modified version, and disposes the second workbook, demonstrating proper resource cleanup with Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook wbCreate = new Workbook();                     // create workbook
        Worksheet wsCreate = wbCreate.Worksheets[0];           // get first worksheet

        // Populate some cells
        wsCreate.Cells["A1"].PutValue("Sample Text");
        wsCreate.Cells["B1"].PutValue(42);

        // Save the newly created workbook
        wbCreate.Save("CreatedWorkbook.xlsx");

        // Explicitly release resources for the created workbook
        wbCreate.Dispose();

        // ---------- Load an existing workbook ----------
        Workbook wbLoad = new Workbook("CreatedWorkbook.xlsx"); // load workbook from file
        Worksheet wsLoad = wbLoad.Worksheets[0];               // get first worksheet

        // Modify the workbook
        wsLoad.Cells["A2"].PutValue(DateTime.Now);
        wsLoad.Cells["B2"].PutValue(3.14159);

        // Save the modified workbook
        wbLoad.Save("ModifiedWorkbook.xlsx");

        // Explicitly release resources for the loaded workbook
        wbLoad.Dispose();
    }
}
