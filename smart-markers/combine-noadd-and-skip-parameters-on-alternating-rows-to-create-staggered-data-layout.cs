// Title: Create a staggered column layout in Excel with Aspose.Cells .NET using ImportObjectArray (skip & start offset)
// Description: This example shows how to build a workbook, define two object arrays, and import each array horizontally while skipping one column between values. The second array starts at column 1, producing an alternating (staggered) column arrangement, then saves the file as StaggeredDataLayout.xlsx.
// Keywords: Aspose.Cells .NET | ImportObjectArray skip | column offset Excel | staggered data layout | C# Excel automation | skip parameter Aspose.Cells | offset rows Aspose.Cells | horizontal import Excel
// Common Searches: Aspose.Cells import object array with column skip | how to offset rows for staggered columns in Excel C# | skip columns while importing data with Aspose.Cells | alternating column layout using ImportObjectArray | combine skip and start column in Aspose.Cells
// Developer Intent: Produce an Excel sheet where each successive row begins one column to the right, using ImportObjectArray’s skip and start‑column arguments to achieve a staggered visual layout.
// Use Cases: Design side‑by‑side comparison tables where each row is shifted to avoid column overlap. | Generate printable schedules or timetables with offset entries for clearer visual separation. | Create multi‑section reports that visually separate sections by inserting blank columns without adding extra rows.
// AI Prompts: Demonstrate how to vary the skip value per row while keeping the staggered effect with ImportObjectArray. | Show an example of vertical ImportObjectArray combined with alternating column offsets. | Explain how the optional noAdd flag works together with skip to insert blank cells in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsStaggeredLayout
{
    // This example shows how to build a workbook, define two object arrays, and import each array horizontally while skipping one column between values. The second array starts at column 1, producing an alternating (staggered) column arrangement, then saves the file as StaggeredDataLayout.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Data for the first row (will be placed in columns A, C, E, ...)
            object[] firstRowData = new object[] { "Name", "Age", "City", "Alice", 30, "New York", "Bob", 25, "Paris" };

            // Data for the second row (will be placed in columns B, D, F, ...)
            object[] secondRowData = new object[] { "Name", "Age", "City", "Charlie", 28, "London", "Diana", 32, "Tokyo" };

            // Import first row data horizontally, skipping one column between entries
            // Parameters: data array, start row 0, start column 0, horizontal (false), skip 1 column
            sheet.Cells.ImportObjectArray(firstRowData, 0, 0, false, 1);

            // Import second row data horizontally, also skipping one column,
            // but start from column 1 to achieve the staggered (offset) layout
            sheet.Cells.ImportObjectArray(secondRowData, 1, 1, false, 1);

            // Save the workbook
            workbook.Save("StaggeredDataLayout.xlsx");
        }
    }
}
