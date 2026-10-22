// Title: Create a staggered Excel column by inserting values into every other row with a skip parameter and stop at a maximum row count using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to write a list of strings into a worksheet column, placing each entry one row apart and halting when a predefined row limit is reached. | Show how to apply a skip interval and a no‑add stop condition in Aspose.Cells to produce a staggered data layout in an Excel file.
// Common Searches: how to write data to every other row in Excel using Aspose.Cells C# | Aspose.Cells example for skipping rows and limiting insertion count | C# create staggered column layout with row limit in Aspose.Cells | insert values with gaps between rows using Aspose.Cells .NET | prevent adding rows after reaching maximum index with Aspose.Cells
// Tags: row skipping insertion Aspose.Cells | row limit enforcement Aspose.Cells | alternating row placement Excel .NET | staggered data layout using Aspose.Cells | conditional row addition C# Aspose.Cells

using Aspose.Cells;
using System;

// // This program creates a new workbook, writes a list of strings into every other row of the first column while respecting a maximum row limit, and saves the result as StaggeredDataLayout.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data to be placed in the worksheet
        string[] values = { "Item1", "Item2", "Item3", "Item4", "Item5", "Item6" };

        // Parameters:
        //   skip   – number of rows to skip between each entry (staggered layout)
        //   noAdd  – stop adding when the target row exceeds a predefined limit
        int skip = 1;               // skip one row between entries
        int maxRows = 12;           // maximum rows allowed (noAdd condition)
        int startRow = 0;           // first row to start inserting
        int columnIndex = 0;        // column where data will be placed

        bool noAdd = false;         // flag indicating whether further addition is prohibited

        for (int i = 0; i < values.Length; i++)
        {
            // Calculate the row index taking the skip into account
            int targetRow = startRow + i * (skip + 1);

            // Apply the noAdd logic: if the target row exceeds the limit, stop inserting
            if (targetRow >= maxRows)
            {
                noAdd = true;
                break;
            }

            // Insert the value into the calculated cell
            cells[targetRow, columnIndex].PutValue(values[i]);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("StaggeredDataLayout.xlsx");
    }
}
