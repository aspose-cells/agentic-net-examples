using System;
using Aspose.Cells;

namespace AsposeCellsSkipExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load a template workbook)
            Workbook workbook = new Workbook(); // new workbook instance
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data to be inserted into the worksheet
            // The data will be placed vertically, one entry per row,
            // with a blank row skipped between each entry.
            object[] data = new object[]
            {
                "First Entry",
                "Second Entry",
                "Third Entry",
                "Fourth Entry"
            };

            // Import the data array starting at row 0, column 0.
            // Parameters:
            //   firstRow   = 0          (start at the first row)
            //   firstColumn= 0          (start at the first column)
            //   isVertical = true       (import vertically)
            //   skip       = 1          (skip one row after each entry)
            sheet.Cells.ImportObjectArray(data, 0, 0, true, 1);

            // Save the workbook to a file
            workbook.Save("AlternatingEntries.xlsx");
        }
    }
}