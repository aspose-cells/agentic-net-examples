// Title: C# – ImportObjectArray with skip=1 to populate alternating rows in Aspose.Cells
// Description: Demonstrates how to create a workbook, import a one‑dimensional array vertically, and use the skip parameter (set to 1) so each entry occupies every other row, then saves the file as AlternatingEntries.xlsx.
// Keywords: Aspose.Cells | ImportObjectArray | skip parameter | alternating rows | C# spreadsheet automation | blank rows between data | vertical import | Excel template population
// Common Searches: Aspose.Cells skip rows example C# | ImportObjectArray every other row | how to leave blank rows when importing data Aspose.Cells | C# populate Excel template with gaps | Aspose.Cells skip parameter usage
// Developer Intent: Insert array values into a worksheet while automatically leaving one empty row between each entry using the skip argument.
// Use Cases: Generate reports where each record is visually separated by a blank line. | Create templates that reserve alternate rows for user comments or signatures. | Design schedules that interleave event rows with empty rows for additional notes.
// AI Prompts: Write C# code that uses Aspose.Cells ImportObjectArray with skip=2 to leave two empty rows between entries. | Explain the effect of the skip parameter in ImportObjectArray and how to calculate the value for a desired row spacing. | Show an example of ImportObjectArray with horizontal orientation and a skip value that inserts empty columns between data items.

using System;
using Aspose.Cells;

namespace AsposeCellsSkipRowsDemo
{
    // Demonstrates how to create a workbook, import a one‑dimensional array vertically, and use the skip parameter (set to 1) so each entry occupies every other row, then saves the file as AlternatingEntries.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data to be inserted into the worksheet
            // Each entry will be placed in a separate row
            object[] data = new object[]
            {
                "Entry 1",
                "Entry 2",
                "Entry 3",
                "Entry 4",
                "Entry 5"
            };

            // Import the data vertically starting at row 0, column 0
            // The 'skip' parameter is set to 1, which means one empty row will be left
            // between each imported entry (i.e., every other row is used)
            sheet.Cells.ImportObjectArray(data, firstRow: 0, firstColumn: 0, isVertical: true, skip: 1);

            // Save the workbook to a file
            workbook.Save("AlternatingEntries.xlsx");
        }
    }
}
