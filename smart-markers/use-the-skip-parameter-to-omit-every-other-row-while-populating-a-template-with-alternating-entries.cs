// Title: Populate an Excel template with alternating rows using Aspose.Cells skip parameter in C#
// Common Searches: C# Aspose.Cells example to write data to every other row in a worksheet | How to use the skip parameter with smart markers to leave blank rows in Excel | Aspose.Cells populate template skipping rows programmatically C# | Insert list items into Excel template with a row gap using Aspose.Cells | Smart markers skip rows while filling template in C# application
// Tags: Aspose.Cells skip rows C# | smart markers alternating rows | populate Excel template with row spacing | C# write data to non‑contiguous rows | Excel template data insertion using Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The program loads a template workbook, iterates a list of strings, and writes each entry into column A using a calculated target row (startRow + i * 2) so that entries are placed in every other row, then saves the populated workbook as PopulatedResult.xlsx.
class Program
{
    static void Main()
    {
        // Load the template workbook
        Workbook workbook = new Workbook("Template.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Example data to insert
        List<string> entries = new List<string>
        {
            "Entry 1",
            "Entry 2",
            "Entry 3",
            "Entry 4",
            "Entry 5"
        };

        // Define the starting row (0‑based). Assuming row 0 is a header, start at row 1.
        int startRow = 1;
        int columnIndex = 0; // Column A

        // Populate every other row (skip one row each time)
        for (int i = 0; i < entries.Count; i++)
        {
            int targetRow = startRow + i * 2; // skip parameter = 2
            Cell cell = sheet.Cells[targetRow, columnIndex];
            cell.PutValue(entries[i]);
        }

        // Save the populated workbook
        workbook.Save("PopulatedResult.xlsx");
    }
}
