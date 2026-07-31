// Title: Staggered Excel Row Layout with Alternating Skip (and NoAdd) Using ImportObjectArray in Aspose.Cells for .NET
// Description: Shows how to import rows into an Excel worksheet with alternating skip values (0 for even rows, 1 for odd rows) via Cells.ImportObjectArray. The sample also explains how to combine the noadd flag with skip to achieve a staggered column layout, then saves the workbook as StaggeredDataLayout.xlsx.
// Keywords: Aspose.Cells | ImportObjectArray | skip parameter | noadd parameter | staggered rows | alternating skip | .NET | C# Excel export | smart markers | data layout | Excel column offset
// Common Searches: Aspose.Cells ImportObjectArray skip example | how to offset every other row in Excel using Aspose.Cells | combine noadd and skip in smart markers | staggered column layout C# Aspose.Cells | alternating column offset Aspose.Cells .NET
// Developer Intent: The developer wants to import data rows into an Excel sheet with alternating skip (and optionally noadd) values to produce a staggered visual layout.
// Use Cases: Create a report where every second row is indented by one column for visual grouping. | Generate a schedule with offset rows to separate time blocks. | Export a list of items where alternate rows are shifted to improve readability in printed Excel sheets.
// AI Prompts: Modify the code to use a custom skip pattern such as 0,2,0,2 while keeping noadd disabled. | Provide an example that imports data vertically with alternating skip values and demonstrates the effect of the noadd flag. | Explain step‑by‑step how to combine the noadd and skip parameters in smart markers to achieve a staggered data layout.

using System;
using Aspose.Cells;

namespace AsposeCellsStaggeredImportDemo
{
    // Shows how to import rows into an Excel worksheet with alternating skip values (0 for even rows, 1 for odd rows) via Cells.ImportObjectArray. The sample also explains how to combine the noadd flag with skip to achieve a staggered column layout, then saves the workbook as StaggeredDataLayout.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data for each row (three columns per row)
            object[][] rowsData = new object[][]
            {
                new object[] { "Item1", 10, DateTime.Today },
                new object[] { "Item2", 20, DateTime.Today.AddDays(1) },
                new object[] { "Item3", 30, DateTime.Today.AddDays(2) },
                new object[] { "Item4", 40, DateTime.Today.AddDays(3) },
                new object[] { "Item5", 50, DateTime.Today.AddDays(4) }
            };

            // Import each row with alternating skip (staggered layout)
            // Even rows: no skip (continuous cells)
            // Odd rows : skip one column between each entry
            for (int rowIndex = 0; rowIndex < rowsData.Length; rowIndex++)
            {
                object[] currentRow = rowsData[rowIndex];
                bool isVertical = false; // import horizontally (row wise)

                // Determine skip value: 0 for even rows, 1 for odd rows
                int skip = (rowIndex % 2 == 0) ? 0 : 1;

                // Import the array starting at column 0 of the current row
                // ImportObjectArray(object[] objArray, int firstRow, int firstColumn, bool isVertical, int skip)
                cells.ImportObjectArray(currentRow, rowIndex, 0, isVertical, skip);
            }

            // Save the workbook to a file
            workbook.Save("StaggeredDataLayout.xlsx");
        }
    }
}
